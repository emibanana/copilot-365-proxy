"""
OAuth2 Device Code & PKCE authentication module for Microsoft 365 Copilot.

Handles the OAuth2 token lifecycle:
  - Device code flow (first-time auth)
  - PKCE authorization code flow (browser-based)
  - Token caching to ~/.hermes/credentials/copilot365_token.json
  - Automatic token refresh

Usage:
    python oauth.py device-code    # Device code flow (headless-friendly)
    python oauth.py pkce           # PKCE flow (opens browser)
"""

import os
import json
import time
import secrets
import hashlib
import base64
import webbrowser
import string
from pathlib import Path
from typing import Optional, Dict, Any

import requests
from dotenv import load_dotenv

load_dotenv()


# --- Configuration ---
AZURE_TENANT_ID = os.getenv("AZURE_TENANT_ID")
AZURE_CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
AZURE_CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")

TOKEN_CACHE_PATH = Path.home() / ".hermes" / "credentials" / "copilot365_token.json"

SCOPES = [
    "https://graph.microsoft.com/Sites.Read.All",
    "https://graph.microsoft.com/Mail.Read",
    "https://graph.microsoft.com/People.Read.All",
    "https://graph.microsoft.com/OnlineMeetingTranscript.Read.All",
    "https://graph.microsoft.com/Chat.Read",
    "https://graph.microsoft.com/ChannelMessage.Read.All",
    "https://graph.microsoft.com/ExternalItem.Read.All",
]


# --- PKCE helpers ---
def _generate_code_verifier(length: int = 64) -> str:
    alphabet = string.ascii_letters + string.digits + "-._~"
    return "".join(secrets.choice(alphabet) for _ in range(length))


def _code_challenge_from_verifier(verifier: str) -> str:
    digest = hashlib.sha256(verifier.encode("utf-8")).digest()
    return base64.urlsafe_b64encode(digest).rstrip(b"=").decode("utf-8")


# --- Token management ---
def _save_token(data: Dict[str, Any]):
    TOKEN_CACHE_PATH.parent.mkdir(parents=True, exist_ok=True)
    data["expires_at"] = time.time() + int(data["expires_in"])
    with open(TOKEN_CACHE_PATH, "w") as f:
        json.dump(data, f, indent=2)
    print(f"Token cached to {TOKEN_CACHE_PATH}")


def get_cached_token() -> Optional[Dict[str, Any]]:
    if TOKEN_CACHE_PATH.exists():
        with open(TOKEN_CACHE_PATH) as f:
            data = json.load(f)
        if data.get("expires_at", 0) > time.time() + 60:
            return data
        if "refresh_token" in data:
            return refresh_token(data["refresh_token"])
    return None


def refresh_token(refresh_token: str) -> Optional[Dict[str, Any]]:
    url = f"https://login.microsoftonline.com/{AZURE_TENANT_ID}/oauth2/v2.0/token"
    payload = {
        "client_id": AZURE_CLIENT_ID,
        "client_secret": AZURE_CLIENT_SECRET,
        "refresh_token": refresh_token,
        "grant_type": "refresh_token",
        "scope": " ".join(SCOPES),
    }
    resp = requests.post(url, data=payload, timeout=10)
    if resp.status_code == 200:
        data = resp.json()
        _save_token(data)
        return data
    print(f"Token refresh failed: {resp.status_code} {resp.text[:200]}")
    return None


# --- OAuth2 flows ---
def device_code_flow():
    """Device code flow - works in headless/SSH environments."""
    print("\n=== Device Code Flow ===")
    print("Starting device code authorization...\n")

    # Start device code flow
    url = f"https://login.microsoftonline.com/{AZURE_TENANT_ID}/oauth2/v2.0/devicecode"
    payload = {
        "client_id": AZURE_CLIENT_ID,
        "scope": " ".join(SCOPES),
    }
    resp = requests.post(url, data=payload, timeout=10)
    if resp.status_code != 200:
        print(f"Failed to start device code flow: {resp.text[:500]}")
        return False

    device_data = resp.json()
    print(f"1. Go to: https://login.microsoft.com/device")
    print(f"2. Enter code: {device_data['user_code']}")
    print(f"   (Or click: {device_data['verification_uri']})")
    print(f"3. Authenticate with your Microsoft 365 account\n")
    print(f"Code expires in {device_data['expires_in']} seconds. Waiting...\n")

    # Poll for token
    poll_url = f"https://login.microsoftonline.com/{AZURE_TENANT_ID}/oauth2/v2.0/token"
    poll_payload = {
        "client_id": AZURE_CLIENT_ID,
        "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
        "device_code": device_data["device_code"],
    }
    timeout = time.time() + device_data["expires_in"]
    while time.time() < timeout:
        time.sleep(device_data.get("interval", 5))
        poll_resp = requests.post(poll_url, data=poll_payload, timeout=10)
        if poll_resp.status_code == 200:
            _save_token(poll_resp.json())
            print("Authorization successful!")
            return True
        error = poll_resp.json().get("error", "")
        if error == "authorization_pending":
            continue
        elif error == "authorization_declined":
            print("Authorization declined by user.")
            return False
        elif error == "expired_token":
            print("Code expired. Restart the flow.")
            return False
        print(f"Poll error: {poll_resp.text[:200]}")

    print("Timed out waiting for authorization.")
    return False


def pkce_flow():
    """PKCE authorization code flow - opens browser for interactive auth."""
    print("\n=== PKCE Authorization Code Flow ===\n")

    verifier = _generate_code_verifier()
    challenge = _code_challenge_from_verifier(verifier)
    state = secrets.token_urlsafe(16)

    params = {
        "client_id": AZURE_CLIENT_ID,
        "response_type": "code",
        "redirect_uri": "http://localhost:8081/callback",
        "response_mode": "query",
        "scope": " ".join(SCOPES),
        "state": state,
        "code_challenge": challenge,
        "code_challenge_method": "S256",
    }
    query = "&".join(f"{k}={requests.utils.quote(v)}" for k, v in params.items())
    auth_url = (
        f"https://login.microsoftonline.com/{AZURE_TENANT_ID}"
        f"/oauth2/v2.0/authorize?{query}"
    )

    print("1. A browser window will open")
    print("2. Sign in and grant permissions")
    print("3. After redirect, the auth code will be exchanged\n")

    try:
        webbrowser.open(auth_url)
    except Exception:
        print(f"Open this URL in your browser:\n{auth_url}\n")

    # Start a simple HTTP server to catch the callback
    import socket
    import urllib.parse

    server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    server.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    server.bind(("127.0.0.1", 8081))
    server.listen(1)
    server.settimeout(300)

    print("Waiting for redirect to http://localhost:8081/callback...")
    try:
        conn, addr = server.accept()
        data = conn.recv(4096).decode("utf-8")
        first_line = data.split("\n")[0]
        path_part = first_line.split(" ")[1]

        # Parse the query params from the callback
        parsed = urllib.parse.urlparse(path_part)
        params = urllib.parse.parse_qs(parsed.query)
        code = params.get("code", [None])[0]
        returned_state = params.get("state", [None])[0]

        if returned_state != state:
            conn.sendall(b"HTTP/1.1 400 Bad Request\r\n\r\nState mismatch")
            conn.close()
            print("Error: state parameter mismatch (possible CSRF)")
            return False

        if not code:
            error = params.get("error", ["unknown"])[0]
            conn.sendall(
                f"HTTP/1.1 400 Bad Request\r\n\r\nOAuth error: {error}".encode()
            )
            conn.close()
            print(f"OAuth error: {error}")
            return False

        # Exchange code for token
        token_url = (
            f"https://login.microsoftonline.com/{AZURE_TENANT_ID}"
            f"/oauth2/v2.0/token"
        )
        payload = {
            "client_id": AZURE_CLIENT_ID,
            "client_secret": AZURE_CLIENT_SECRET,
            "grant_type": "authorization_code",
            "code": code,
            "redirect_uri": "http://localhost:8081/callback",
            "code_verifier": verifier,
            "scope": " ".join(SCOPES),
        }
        token_resp = requests.post(token_url, data=payload, timeout=10)
        if token_resp.status_code != 200:
            conn.sendall(b"HTTP/1.1 500 Error\r\n\r\nToken exchange failed")
            conn.close()
            print(f"Token exchange failed: {token_resp.text[:500]}")
            return False

        _save_token(token_resp.json())
        conn.sendall(
            b"HTTP/1.1 200 OK\r\nContent-Type: text/html\r\n\r\n"
            b"<html><body><h2>Authorization complete!</h2>"
            b"<p>You may close this window and return to the terminal.</p></body></html>"
        )
        conn.close()
        print("\nAuthorization successful!")
        return True

    except socket.timeout:
        print("Timed out waiting for browser redirect.")
        return False
    finally:
        server.close()


if __name__ == "__main__":
    import sys

    if not all([AZURE_TENANT_ID, AZURE_CLIENT_ID]):
        print("ERROR: Missing Azure credentials.")
        print("Ensure .env has AZURE_TENANT_ID and AZURE_CLIENT_ID.")
        sys.exit(1)

    cached = get_cached_token()
    if cached:
        print(f"Valid cached token found. Expires at: {time.ctime(cached['expires_at'])}")
        print("Use --force to re-authorize.")
        if "--force" not in sys.argv:
            sys.exit(0)

    method = sys.argv[1] if len(sys.argv) > 1 else "device-code"
    if method == "device-code" or method == "device_code":
        success = device_code_flow()
    elif method == "pkce":
        success = pkce_flow()
    else:
        print(f"Unknown method: {method}")
        print("Usage: python oauth.py [device-code|pkce] [--force]")
        sys.exit(1)

    sys.exit(0 if success else 1)
