import json
import time
import os
import base64
import requests
import uvicorn
from fastapi import FastAPI, HTTPException
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from typing import List, Optional
from dotenv import load_dotenv

# Load .env from the same directory
load_dotenv()

GRAPH_VERSION = os.getenv("GRAPH_API_VERSION", "beta")

# In-memory conversation store: "{user_oid}:{model}" -> {"conversation_id": str, "turn_count": int}
_user_conversations: dict[str, dict] = {}

# Cache for discovered Copilot agents: refreshed on demand
_agents_cache: dict = {"data": None, "fetched_at": 0}
_AGENTS_CACHE_TTL = 300  # seconds

DEFAULT_MODEL_ID = "copilot-chat"

# Max turns before rotating to a fresh conversation (avoids token-limit issues)
MAX_TURNS_PER_CONVERSATION = int(os.getenv("COPILOT_MAX_TURNS", "50"))


def _get_user_oid(access_token: str) -> str:
    try:
        parts = access_token.split(".")
        if len(parts) == 3:
            payload = parts[1] + "=="
            decoded = base64.urlsafe_b64decode(payload)
            claims = json.loads(decoded)
            oid = claims.get("oid")
            if oid:
                return oid
    except Exception:
        pass
    return "default"


def _load_token():
    """Load the cached OAuth2 token from disk."""
    cred_path = os.path.expanduser("~/.hermes/credentials/copilot365_token.json")
    try:
        with open(cred_path) as f:
            data = json.load(f)
        access_token = data.get("access_token", "")
        if not access_token:
            raise HTTPException(401, detail="No access token")
        return access_token
    except Exception as e:
        raise HTTPException(500, detail=f"Token error: {e}") from e


def _discover_copilot_agents(access_token: str) -> list[dict]:
    """Fetch available Copilot agents/extensions from Graph API with caching."""
    now = time.time()
    if _agents_cache["data"] is not None and now - _agents_cache["fetched_at"] < _AGENTS_CACHE_TTL:
        return _agents_cache["data"]

    headers = {"Authorization": f"Bearer {access_token}"}
    resp = requests.get(
        f"https://graph.microsoft.com/{GRAPH_VERSION}/copilot/extensions",
        headers=headers,
        timeout=15,
    )
    if resp.status_code == 200:
        agents = resp.json().get("value", [])
        _agents_cache["data"] = agents
        _agents_cache["fetched_at"] = now
        return agents

    # Endpoint unavailable (not all tenants have it) — return empty
    _agents_cache["data"] = []
    _agents_cache["fetched_at"] = now
    return []


def _agent_id_from_model(model: str) -> Optional[str]:
    """Extract the raw agent/extension ID from a model string like 'copilot-agent-<id>'."""
    if model.startswith("copilot-agent-"):
        return model[len("copilot-agent-"):]
    return None


def _create_conversation(access_token: str, model: str = DEFAULT_MODEL_ID) -> dict:
    """Create a new Copilot conversation, optionally binding it to an agent."""
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }
    body: dict = {}
    agent_id = _agent_id_from_model(model)
    if agent_id:
        body["copilotAgent"] = {
            "agentDefinition": {
                "extensions": [
                    {
                        "@odata.type": "#microsoft.graph.aiPlugin",
                        "extensionId": agent_id,
                    }
                ]
            }
        }
    create = requests.post(
        f"https://graph.microsoft.com/{GRAPH_VERSION}/copilot/conversations",
        headers=headers,
        json=body,
        timeout=30,
    )
    if create.status_code not in (200, 201):
        raise HTTPException(create.status_code, detail=create.text)
    conv = create.json()
    conversation_id = conv.get("id")
    if not conversation_id:
        raise HTTPException(500, detail="No conversation ID")
    return {"conversation_id": conversation_id, "turn_count": 0, "user_oid": _get_user_oid(access_token)}


def _call_copilot(access_token: str, user_message: str, model: str = DEFAULT_MODEL_ID) -> dict:
    """Call the Microsoft Graph Copilot API and return the response data."""
    user_oid = _get_user_oid(access_token)
    conv_key = f"{user_oid}:{model}"

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }

    # Get or create a per-(user, model) conversation
    conv_data = _user_conversations.get(conv_key)
    if not conv_data:
        conv_data = _create_conversation(access_token, model)
        conversation_id = conv_data["conversation_id"]
        _user_conversations[conv_key] = conv_data
    else:
        conversation_id = conv_data["conversation_id"]
        # Rotate if too many turns (prevents hitting Copilot's context window limit)
        if conv_data["turn_count"] >= MAX_TURNS_PER_CONVERSATION:
            conv_data = _create_conversation(access_token, model)
            conversation_id = conv_data["conversation_id"]
            _user_conversations[conv_key] = conv_data

    # Send the chat message
    tz = os.getenv("USER_TIMEZONE", "UTC")
    chat_payload = {
        "message": {"text": user_message},
        "locationHint": {"timeZone": tz},
    }

    chat_resp = requests.post(
        f"https://graph.microsoft.com/{GRAPH_VERSION}/copilot/conversations/{conversation_id}/chat",
        headers=headers,
        json=chat_payload,
        timeout=60,
    )

    # Handle stale conversations (404/410/400 -> recreate)
    if chat_resp.status_code not in (200, 201):
        if chat_resp.status_code in (404, 410, 400):
            _user_conversations.pop(conv_key, None)
            conv_data2 = _create_conversation(access_token, model)
            conversation_id = conv_data2["conversation_id"]
            _user_conversations[conv_key] = conv_data2
            chat_resp = requests.post(
                f"https://graph.microsoft.com/{GRAPH_VERSION}/copilot/conversations/{conversation_id}/chat",
                headers=headers,
                json=chat_payload,
                timeout=60,
            )
        if chat_resp.status_code not in (200, 201):
            raise HTTPException(chat_resp.status_code, detail=chat_resp.text)

    # Increment turn counter
    conv_entry = _user_conversations.get(conv_key)
    if conv_entry and conv_entry.get("conversation_id") == conversation_id:
        conv_entry["turn_count"] = conv_entry.get("turn_count", 0) + 1

    return chat_resp.json()


def _build_openai_response(graph_data: dict, model: str) -> dict:
    """Map Graph Copilot response to OpenAI chat.completion format."""
    msgs = graph_data.get("messages", [])
    assistant_content = msgs[-1].get("text", "(no response)") if msgs else "(no response)"
    return {
        "id": graph_data.get("id", ""),
        "object": "chat.completion",
        "created": int(time.time()),
        "model": model,
        "choices": [
            {
                "index": 0,
                "message": {"role": "assistant", "content": assistant_content},
                "finish_reason": "stop",
            }
        ],
    }


class Message(BaseModel):
    role: str
    content: str


class ChatCompletionRequest(BaseModel):
    model: str
    messages: List[Message]
    stream: bool = False


app = FastAPI(title="Copilot365 Proxy")


@app.get("/")
def root():
    return {"status": "ok", "service": "copilot365-proxy"}


@app.get("/v1/models")
def list_models():
    """OpenAI-compatible model list — includes default Copilot plus any discovered agents."""
    now = int(time.time())
    models = [
        {
            "id": DEFAULT_MODEL_ID,
            "object": "model",
            "created": now,
            "owned_by": "microsoft-365",
        }
    ]
    try:
        access_token = _load_token()
        agents = _discover_copilot_agents(access_token)
        for agent in agents:
            agent_id = agent.get("id") or agent.get("extensionId") or agent.get("pluginId")
            if not agent_id:
                continue
            display_name = agent.get("displayName") or agent.get("name") or agent_id
            models.append(
                {
                    "id": f"copilot-agent-{agent_id}",
                    "object": "model",
                    "created": now,
                    "owned_by": "microsoft-365",
                    "display_name": display_name,
                }
            )
    except Exception:
        pass  # Token not yet available or endpoint unsupported — return defaults only
    return {"object": "list", "data": models}


@app.post("/v1/chat/completions")
async def chat_completions(request: ChatCompletionRequest):
    access_token = _load_token()

    # Extract the last user message
    user_message = None
    for msg in reversed(request.messages):
        if msg.role == "user":
            user_message = msg.content
            break
    if user_message is None:
        raise HTTPException(400, detail="No user message")

    if request.stream:
        return _handle_streaming(access_token, user_message, request.model)
    else:
        return _handle_non_streaming(access_token, user_message, request.model)


def _handle_non_streaming(access_token: str, user_message: str, model: str):
    """Non-streaming response — return full JSON."""
    graph_data = _call_copilot(access_token, user_message, model)
    return _build_openai_response(graph_data, model)


def _handle_streaming(access_token: str, user_message: str, model: str):
    """Streaming response — return SSE chunks.

    The Graph Copilot API does not support streaming, so we fake it by
    sending the full response as a single content chunk plus a finish chunk.
    Most AI agent frameworks (Hermes, OpenClaw, etc.) use streaming by default,
    so this is required for compatibility.
    """
    graph_data = _call_copilot(access_token, user_message, model)
    response_data = _build_openai_response(graph_data, model)

    response_id = response_data["id"]
    created = response_data["created"]
    content = response_data["choices"][0]["message"]["content"]

    async def generate():
        # Content chunk
        chunk = {
            "id": response_id,
            "object": "chat.completion.chunk",
            "created": created,
            "model": model,
            "choices": [
                {
                    "index": 0,
                    "delta": {"role": "assistant", "content": content},
                    "finish_reason": None,
                }
            ],
        }
        yield f"data: {json.dumps(chunk)}\n\n"

        # Final chunk with finish_reason
        finish_chunk = {
            "id": response_id,
            "object": "chat.completion.chunk",
            "created": created,
            "model": model,
            "choices": [
                {
                    "index": 0,
                    "delta": {},
                    "finish_reason": "stop",
                }
            ],
        }
        yield f"data: {json.dumps(finish_chunk)}\n\n"

        yield "data: [DONE]\n\n"

    return StreamingResponse(
        generate(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",
        },
    )


if __name__ == "__main__":
    host = os.getenv("COPILOT_PROXY_HOST", "127.0.0.1")
    port = int(os.getenv("COPILOT_PROXY_PORT", "8081"))
    uvicorn.run(app, host=host, port=port, log_level="info")
