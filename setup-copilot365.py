#!/usr/bin/env python3
"""
Interactive wizard to configure Microsoft 365 Copilot proxy.
"""

import os
from pathlib import Path

CONFIG_PATH = Path.home() / ".hermes" / "services" / "copilot365-proxy" / ".env"

print("=" * 60)
print("Microsoft 365 Copilot Proxy - Setup Wizard")
print("=" * 60)

print("""
This wizard sets up the Copilot365 proxy, which exposes an OpenAI-compatible
HTTP API backed by Microsoft 365 Copilot's Graph API.

PREREQUISITES (Azure AD):
  1. App registration (confidential client) in your Azure AD tenant
  2. Authentication:
     - Add platform: "Mobile and desktop applications" or "Web"
     - "Allow public client flows" must be OFF
  3. API permissions -> Add Microsoft Graph -> Delegated permissions:
     - Sites.Read.All
     - Mail.Read
     - People.Read.All
     - OnlineMeetingTranscript.Read.All
     - Chat.Read
     - ChannelMessage.Read.All
     - ExternalItem.Read.All
     (There is NO "Copilot.Chat" permission -- these 7 are required)
  4. Grant admin consent for all permissions
  5. Certificates & secrets -> create a client secret -> copy the VALUE
""")

tenant_id = input("Azure Tenant ID: ").strip()
client_id = input("Azure Client (Application) ID: ").strip()
client_secret = input("Azure Client Secret (value, not ID): ").strip()

graph_version = input("Graph API version [beta]: ").strip() or "beta"
timezone = input("User timezone for Copilot [UTC]: ").strip() or "UTC"

# Write .env
env_lines = [
    f"AZURE_TENANT_ID={tenant_id}",
    f"AZURE_CLIENT_ID={client_id}",
    f"AZURE_CLIENT_SECRET={client_secret}",
    f"GRAPH_API_VERSION={graph_version}",
    f"USER_TIMEZONE={timezone}",
    "",
    "# Optional: bind address",
    "COPILOT_PROXY_HOST=127.0.0.1",
    "COPILOT_PROXY_PORT=8081",
]

CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
CONFIG_PATH.write_text("\n".join(env_lines) + "\n")
print(f"\nConfiguration saved to {CONFIG_PATH}")
print("\nNext steps:")
print("  1. Verify Azure app is configured as confidential client")
print("  2. Start proxy: copilot365-proxy start")
print("  3. Make a request to trigger device code flow")
print("  4. Visit https://login.microsoft.com/device and enter the code")
print("  5. Token will be cached at ~/.hermes/credentials/copilot365_token.json")
