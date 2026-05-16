"""
Microsoft 365 Copilot → OpenAI-compatible proxy server.

Translates Hermes/OpenAI /v1/chat/completions requests into Microsoft Graph
Copilot Chat API calls with automatic conversation lifecycle management.

Architecture:
  Hermes Agent → HTTP POST /v1/chat/completions
       ↓
  Copilot Proxy (localhost:8081)
       ↓ OAuth2 Bearer + /beta/copilot/conversations/{id}/chat
  Microsoft Graph API
"""

import json
import time
import os
import base64
import requests
import uvicorn
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from typing import List
from dotenv import load_dotenv

# Load config from .env in same directory
load_dotenv()

GRAPH_VERSION = os.getenv("GRAPH_API_VERSION", "beta")

# In-memory conversation store: user OID -> conversation_id
_user_conversations: dict[str, str] = {}


def _get_user_oid(access_token: str) -> str:
    """Extract the OID claim from the JWT access token for conversation keying."""
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


async def _ensure_authorized():
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


class Message(BaseModel):
    role: str
    content: str


class ChatCompletionRequest(BaseModel):
    model: str
    messages: List[Message]
    stream: bool = False


class CompletionChoice(BaseModel):
    index: int
    message: Message
    finish_reason: str = "stop"


class ChatCompletionResponse(BaseModel):
    id: str
    object: str = "chat.completion"
    created: int
    model: str
    choices: List[CompletionChoice]


app = FastAPI(title="Copilot365 Proxy")


@app.get("/")
def root():
    return {"status": "ok", "service": "copilot365-proxy"}


@app.get("/v1/models")
def list_models():
    """OpenAI-compatible model list endpoint."""
    return {
        "object": "list",
        "data": [
            {
                "id": "copilot-chat",
                "object": "model",
                "created": int(time.time()),
                "owned_by": "microsoft-365",
            }
        ],
    }


@app.post("/v1/chat/completions", response_model=ChatCompletionResponse)
async def chat_completions(request: ChatCompletionRequest):
    access_token = await _ensure_authorized()
    user_oid = _get_user_oid(access_token)

    # Extract the most recent user message
    user_message = None
    for msg in reversed(request.messages):
        if msg.role == "user":
            user_message = msg.content
            break
    if user_message is None:
        raise HTTPException(400, detail="No user message found in request")

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }

    # Get or create a conversation
    conversation_id = _user_conversations.get(user_oid)
    if not conversation_id:
        create = requests.post(
            f"https://graph.microsoft.com/{GRAPH_VERSION}/copilot/conversations",
            headers=headers,
            json={},
            timeout=30,
        )
        if create.status_code not in (200, 201):
            raise HTTPException(create.status_code, detail=create.text)
        conv = create.json()
        conversation_id = conv.get("id")
        if not conversation_id:
            raise HTTPException(500, detail="No conversation ID from Copilot API")
        _user_conversations[user_oid] = conversation_id

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

    # Handle stale conversations (404/410/400 → recreate)
    if chat_resp.status_code not in (200, 201):
        if chat_resp.status_code in (404, 410, 400):
            _user_conversations.pop(user_oid, None)
            create2 = requests.post(
                f"https://graph.microsoft.com/{GRAPH_VERSION}/copilot/conversations",
                headers=headers,
                json={},
                timeout=30,
            )
            if create2.status_code in (200, 201):
                conv2 = create2.json()
                new_id = conv2.get("id")
                if new_id:
                    _user_conversations[user_oid] = new_id
                    conversation_id = new_id
                    chat_resp = requests.post(
                        f"https://graph.microsoft.com/{GRAPH_VERSION}/copilot/conversations/{conversation_id}/chat",
                        headers=headers,
                        json=chat_payload,
                        timeout=60,
                    )
        if chat_resp.status_code not in (200, 201):
            raise HTTPException(chat_resp.status_code, detail=chat_resp.text)

    # Map Graph response to OpenAI format
    data = chat_resp.json()
    msgs = data.get("messages", [])
    assistant_content = msgs[-1].get("text", "(no response)") if msgs else "(no response)"

    return {
        "id": data.get("id", ""),
        "object": "chat.completion",
        "created": int(time.time()),
        "model": request.model,
        "choices": [
            {
                "index": 0,
                "message": {"role": "assistant", "content": assistant_content},
                "finish_reason": "stop",
            }
        ],
    }


if __name__ == "__main__":
    host = os.getenv("COPILOT_PROXY_HOST", "127.0.0.1")
    port = int(os.getenv("COPILOT_PROXY_PORT", "8081"))
    uvicorn.run(app, host=host, port=port, log_level="info")
