"""
One-time (and weekly) Schwab login.

Run on the homelab machine:  python auth_setup.py
It prints a link. Open it in any browser, log in to Schwab, approve the app.
Schwab then redirects to your callback URL, which shows a "can't connect"
error page. That's expected. Copy the FULL URL from the address bar and
paste it back here. A token file is saved and the bot can use it for 7 days.
"""
import os

from dotenv import load_dotenv
from schwab import auth

load_dotenv()

token_path = os.getenv("SCHWAB_TOKEN_PATH", "schwab_token.json")
if os.path.exists(token_path):
    os.remove(token_path)  # start fresh so the 7-day clock resets

client = auth.client_from_manual_flow(
    os.environ["SCHWAB_APP_KEY"],
    os.environ["SCHWAB_APP_SECRET"],
    os.environ["SCHWAB_CALLBACK_URL"],
    token_path,
)
r = client.get_account_numbers()
r.raise_for_status()
accts = ", ".join("..." + a["accountNumber"][-4:] for a in r.json())
print(f"Success. Linked accounts: {accts}")
print(f"Token saved to {token_path}. Good for 7 days.")

# weekly backup rides along with the weekly login
try:
    import backup
    backup.main()
except Exception as e:  # a backup problem should never block the login
    print(f"Login succeeded, but the backup failed: {e}")
