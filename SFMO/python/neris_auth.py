"""Shared NERIS client-credentials login (no email, password, or MFA prompt).

Credentials come from the NERIS_CLIENT_ID and NERIS_CLIENT_SECRET environment
variables, or from a git-ignored .env file in this folder or any parent folder.
"""

import os
import sys
from pathlib import Path

BASE_URL = "https://api.neris.fsri.org/v1"


def _load_dotenv():
    for folder in Path(__file__).resolve().parents:
        env_file = folder / ".env"
        if not env_file.is_file():
            continue
        for line in env_file.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            os.environ.setdefault(key.strip(), value.strip().strip("\"'"))
        return


def authenticate_with_password():
    """Interactive email/password login (the NERIS client prompts for the MFA code)."""
    import getpass
    from neris_api_client import Config, NerisApiClient

    username = input("NERIS Email: ").strip()
    password = getpass.getpass("NERIS Password: ").strip()
    if not username or not password:
        sys.exit("Email and password are required.")

    print("\nConnecting to NERIS API...")
    client = NerisApiClient(
        Config(
            base_url=BASE_URL,
            grant_type="password",
            username=username,
            password=password,
        )
    )
    print("\nCHECK YOUR EMAIL FOR THE MFA CODE")
    client.list_incidents(page_size=1)
    print("✓ Authentication successful!")
    return client


def authenticate_with_fallback(verify_ssl=True):
    """Try the client credentials first; prompt for email/password if missing or rejected."""
    _load_dotenv()
    if not (os.environ.get("NERIS_CLIENT_ID") and os.environ.get("NERIS_CLIENT_SECRET")):
        print("Client credentials not found in environment or .env; falling back to email/password login.")
        return authenticate_with_password()
    try:
        return authenticate_client_credentials(verify_ssl=verify_ssl)
    except Exception as exc:
        print(f"Client credential login failed ({exc}); falling back to email/password login.")
        return authenticate_with_password()


def authenticate_client_credentials(verify_ssl=True):
    from neris_api_client import Config, NerisApiClient

    _load_dotenv()
    client_id = os.environ.get("NERIS_CLIENT_ID")
    client_secret = os.environ.get("NERIS_CLIENT_SECRET")
    if not client_id or not client_secret:
        sys.exit(
            "Missing NERIS_CLIENT_ID / NERIS_CLIENT_SECRET. Set them as environment "
            "variables or in a .env file (see .env.example)."
        )

    if not verify_ssl:
        import urllib3
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    print("\nConnecting to NERIS API...")
    client = NerisApiClient(
        Config(
            base_url=BASE_URL,
            grant_type="client_credentials",
            client_id=client_id,
            client_secret=client_secret,
        )
    )
    if not verify_ssl:
        client._session.verify = False
    client.list_incidents(page_size=1)
    print("✓ Authentication successful!")
    return client
