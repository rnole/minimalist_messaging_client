from __future__ import annotations

import os
import threading
import tkinter as tk
from datetime import datetime, timedelta, timezone
from pathlib import Path
from dataclasses import dataclass
from tkinter import messagebox, ttk
from typing import Callable
from urllib.parse import quote
from dotenv import load_dotenv
import msal
import requests
from requests import HTTPError

load_dotenv()

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
SCOPES = ["Mail.Read", "Mail.Send", "User.Read"]
PERSONAL_ACCOUNT_TENANT = "consumers"
APP_DIR = os.path.dirname(os.path.abspath(__file__))
APPROVED_SENDERS_FILE = os.path.join(APP_DIR, "approved_emails.txt")
TOKEN_CACHE_FILE = os.path.join(APP_DIR, "msal_token_cache.bin")


@dataclass
class EmailMessage:
    message_id: str
    sender: str
    subject: str
    received: str
    body_preview: str


class GraphClient:
    def __init__(self, client_id: str, tenant_id: str) -> None:
        authority = f"https://login.microsoftonline.com/{tenant_id}"
        self._token_cache = msal.SerializableTokenCache()
        if os.path.exists(TOKEN_CACHE_FILE):
            try:
                self._token_cache.deserialize(Path(TOKEN_CACHE_FILE).read_text(encoding="utf-8"))
            except Exception:  # noqa: BLE001
                pass

        self._app = msal.PublicClientApplication(
            client_id=client_id,
            authority=authority,
            token_cache=self._token_cache,
        )
        self._access_token: str | None = None

    def _save_cache(self) -> None:
        if self._token_cache.has_state_changed:
            Path(TOKEN_CACHE_FILE).write_text(self._token_cache.serialize(), encoding="utf-8")

    def authenticate(self) -> bool:
        accounts = self._app.get_accounts()
        if not accounts:
            return False

        result = self._app.acquire_token_silent(SCOPES, account=accounts[0])
        if not result:
            return False

        token = result.get("access_token")
        if not token:
            return False

        self._access_token = token
        self._save_cache()
        return True

    def complete_device_login(self, prompt_callback: Callable[[str], None] | None = None) -> None:
        flow = self._app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise RuntimeError("Could not initiate device flow.")

        if prompt_callback is not None:
            prompt_callback(flow["message"])

        result = self._app.acquire_token_by_device_flow(flow)
        token = result.get("access_token")
        if not token:
            raise RuntimeError(result.get("error_description", "Unable to authenticate."))

        granted_scopes = set(result.get("scope", "").split())
        missing_scopes = [scope for scope in SCOPES if scope not in granted_scopes]
        if missing_scopes:
            raise RuntimeError(
                "Sign-in completed but Microsoft did not grant all required permissions. "
                f"Missing: {', '.join(missing_scopes)}. Sign in again and accept every requested permission."
            )

        self._access_token = token
        self._save_cache()

    def _headers(self) -> dict[str, str]:
        if not self._access_token:
            raise RuntimeError("Not authenticated yet.")
        return {
            "Authorization": f"Bearer {self._access_token}",
            "Content-Type": "application/json",
        }

    @staticmethod
    def _raise_graph_error(exc: HTTPError) -> None:
        response = exc.response
        if response is None:
            raise exc

        status = response.status_code
        reason = (response.reason or "").strip()
        body_text = (response.text or "").strip()
        code = "UnknownError"
        message = ""

        try:
            payload = response.json()
        except ValueError:
            payload = None

        if isinstance(payload, dict):
            err = payload.get("error", {})
            if isinstance(err, dict):
                code = err.get("code") or code
                message = (err.get("message") or "").strip()
            elif isinstance(err, str):
                code = err

            if not message:
                message = (payload.get("error_description") or "").strip()

        if not message:
            message = body_text or reason or "No error details returned by Microsoft Graph."

        request_id = (
            response.headers.get("request-id")
            or response.headers.get("client-request-id")
            or ""
        )
        details = f"Microsoft Graph error ({status} / {code}): {message}"
        if status in {401, 403}:
            details += (
                " Verify delegated Graph scopes (Mail.Read/Mail.Send/User.Read), consent status, and that "
                "you signed in with the mailbox account in the correct tenant. For personal accounts, sign in "
                "again and approve all requested permissions in the device-code browser page."
            )
        if request_id:
            details += f" Request ID: {request_id}."

        raise RuntimeError(details) from exc

    def fetch_messages(self, senders: list[str], max_items: int = 30, lookback_days: int = 2) -> list[EmailMessage]:
        if not senders:
            raise ValueError("At least one sender email is required.")

        allowed = {s.strip().lower() for s in senders if s.strip()}
        cutoff = datetime.now(timezone.utc) - timedelta(days=lookback_days)

        select = "id,subject,receivedDateTime,from,bodyPreview,isRead"
        # Graph can return 400/InefficientFilter on nested sender filters with order-by.
        # To keep behavior reliable, fetch recent inbox messages and apply sender/unread/date
        # constraints locally.
        endpoint = (
            f"{GRAPH_BASE_URL}/me/mailFolders/Inbox/messages"
            f"?$top={max(max_items * 8, 200)}"
            f"&$orderby=receivedDateTime desc"
            f"&$select={quote(select)}"
        )

        response = requests.get(endpoint, headers=self._headers(), timeout=20)
        try:
            response.raise_for_status()
        except HTTPError as exc:
            self._raise_graph_error(exc)

        payload = response.json()
        items = []
        for raw in payload.get("value", []):
            sender = raw.get("from", {}).get("emailAddress", {}).get("address", "Unknown")
            if sender.lower() not in allowed:
                continue

            if raw.get("isRead", True):
                continue

            received_raw = raw.get("receivedDateTime", "")
            try:
                received_dt = datetime.fromisoformat(received_raw.replace("Z", "+00:00"))
            except ValueError:
                continue

            if received_dt < cutoff:
                continue

            items.append(
                EmailMessage(
                    message_id=raw["id"],
                    sender=sender,
                    subject=raw.get("subject") or "(No subject)",
                    received=received_raw,
                    body_preview=raw.get("bodyPreview", ""),
                )
            )
            if len(items) >= max_items:
                break

        return items

    def send_email(self, recipient: str, subject: str, body: str) -> None:
        payload = {
            "message": {
                "subject": subject,
                "body": {"contentType": "Text", "content": body},
                "toRecipients": [
                    {"emailAddress": {"address": recipient}},
                ],
            },
            "saveToSentItems": True,
        }

        response = requests.post(
            f"{GRAPH_BASE_URL}/me/sendMail",
            headers=self._headers(),
            json=payload,
            timeout=20,
        )
        try:
            response.raise_for_status()
        except HTTPError as exc:
            self._raise_graph_error(exc)


class FocusMailApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Focused Outlook Mail")
        self.root.geometry("980x650")
        self._start_maximized()

        self.client: GraphClient | None = None
        self.allowed_senders: list[str] = []
        self.messages: list[EmailMessage] = []

        self._build_ui()
        self._preload_approved_senders()
        self._attempt_silent_sign_in()

    def _start_maximized(self) -> None:
        """Open the desktop app maximized across platforms when possible."""
        try:
            self.root.state("zoomed")  # Windows
            return
        except tk.TclError:
            pass

        try:
            self.root.attributes("-zoomed", True)  # Some Linux window managers
            return
        except tk.TclError:
            pass

        # Fallback: full-screen; user can typically press Esc/F11 depending on OS.
        self.root.attributes("-fullscreen", True)

    def _build_ui(self) -> None:
        frame = ttk.Frame(self.root, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)

        creds = ttk.LabelFrame(frame, text="Microsoft App Registration", padding=10)
        creds.pack(fill=tk.X)

        ttk.Label(creds, text="Tenant ID / authority").grid(row=0, column=0, sticky=tk.W)
        self.tenant_var = tk.StringVar(
            value=os.getenv("OUTLOOK_TENANT_ID", PERSONAL_ACCOUNT_TENANT)
        )
        ttk.Entry(creds, textvariable=self.tenant_var, width=52).grid(row=0, column=1, padx=6)
        ttk.Label(
            creds,
            text=(
                "Use 'consumers' for free personal Outlook accounts, or your Entra tenant ID for work/school."
            ),
        ).grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=(4, 0))

        ttk.Label(creds, text="Client ID").grid(row=1, column=0, sticky=tk.W)
        self.client_var = tk.StringVar(value=os.getenv("OUTLOOK_CLIENT_ID", ""))
        ttk.Entry(creds, textvariable=self.client_var, width=52).grid(row=1, column=1, padx=6)

        ttk.Button(creds, text="Sign in", command=self.sign_in).grid(row=0, column=2, rowspan=2, padx=6)

        senders_frame = ttk.LabelFrame(frame, text="Mandatory senders filter", padding=10)
        senders_frame.pack(fill=tk.X, pady=(10, 0))

        ttk.Label(
            senders_frame,
            text="Enter sender email(s). Only UNREAD emails from these senders in the last 2 days are shown.",
        ).pack(anchor=tk.W)
        self.senders_var = tk.StringVar()
        ttk.Entry(senders_frame, textvariable=self.senders_var).pack(fill=tk.X, pady=4)
        ttk.Button(senders_frame, text="Load emails", command=self.load_emails).pack(anchor=tk.E)

        mailbox_frame = ttk.PanedWindow(frame, orient=tk.HORIZONTAL)
        mailbox_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))

        left = ttk.Frame(mailbox_frame)
        right = ttk.Frame(mailbox_frame)
        mailbox_frame.add(left, weight=2)
        mailbox_frame.add(right, weight=3)

        self.tree = ttk.Treeview(left, columns=("sender", "subject", "received"), show="headings", height=14)
        for col, text, width in (("sender", "Sender", 180), ("subject", "Subject", 280), ("received", "Received", 160)):
            self.tree.heading(col, text=text)
            self.tree.column(col, width=width, anchor=tk.W)
        self.tree.pack(fill=tk.BOTH, expand=True)
        self.tree.bind("<<TreeviewSelect>>", self.on_select_message)

        ttk.Label(right, text="Preview").pack(anchor=tk.W)
        self.preview = tk.Text(right, height=12, wrap=tk.WORD)
        self.preview.pack(fill=tk.BOTH, expand=True)

        ttk.Label(right, text="Reply subject").pack(anchor=tk.W, pady=(8, 0))
        self.reply_subject = tk.StringVar()
        ttk.Entry(right, textvariable=self.reply_subject).pack(fill=tk.X)

        ttk.Label(right, text="Reply body").pack(anchor=tk.W, pady=(6, 0))
        self.reply_body = tk.Text(right, height=8, wrap=tk.WORD)
        self.reply_body.pack(fill=tk.BOTH, expand=True)

        ttk.Button(right, text="Send reply", command=self.send_reply).pack(anchor=tk.E, pady=(6, 0))

        self.status_var = tk.StringVar(value="Please sign in first.")
        ttk.Label(frame, textvariable=self.status_var).pack(anchor=tk.W, pady=(8, 0))

    def _preload_approved_senders(self) -> None:
        if not os.path.exists(APPROVED_SENDERS_FILE):
            return

        try:
            with open(APPROVED_SENDERS_FILE, encoding="utf-8") as infile:
                senders = [
                    line.strip()
                    for line in infile
                    if line.strip() and not line.strip().startswith("#")
                ]
        except OSError:
            return

        if senders:
            self.senders_var.set(", ".join(senders))
            self.status_var.set(
                f"Loaded {len(senders)} approved sender(s) from approved_emails.txt."
            )

    def _attempt_silent_sign_in(self) -> None:
        tenant = self.tenant_var.get().strip() or PERSONAL_ACCOUNT_TENANT
        client = self.client_var.get().strip()
        if not client:
            return

        self.client = GraphClient(client_id=client, tenant_id=tenant)
        if self.client.authenticate():
            self.status_var.set("Signed in from saved session.")

    def sign_in(self) -> None:
        tenant = self.tenant_var.get().strip()
        client = self.client_var.get().strip()

        if not client:
            messagebox.showerror("Missing info", "Client ID is required.")
            return

        if not tenant:
            tenant = PERSONAL_ACCOUNT_TENANT
            self.tenant_var.set(tenant)

        self.client = GraphClient(client_id=client, tenant_id=tenant)
        self.status_var.set("Trying saved session...")

        def _auth() -> None:
            try:
                if self.client.authenticate():
                    self.root.after(0, lambda: self.status_var.set("Signed in from saved session."))
                    return

                self.root.after(0, lambda: self.status_var.set("Signing in with device code flow..."))
                self.client.complete_device_login(
                    prompt_callback=lambda msg: self.root.after(0, lambda: messagebox.showinfo("Authenticate", msg))
                )
                self.root.after(0, lambda: self.status_var.set("Signed in successfully."))
            except Exception as exc:  # noqa: BLE001
                self.root.after(0, lambda: messagebox.showerror("Authentication error", str(exc)))
                self.root.after(0, lambda: self.status_var.set("Sign-in failed."))

        threading.Thread(target=_auth, daemon=True).start()

    def _parse_senders(self) -> list[str]:
        senders = [s.strip().lower() for s in self.senders_var.get().split(",") if s.strip()]
        if not senders:
            raise ValueError("You must enter at least one sender email address.")
        return senders

    def load_emails(self) -> None:
        if not self.client:
            messagebox.showerror("Not signed in", "Sign in before loading emails.")
            return

        try:
            self.allowed_senders = self._parse_senders()
        except ValueError as exc:
            messagebox.showerror("Invalid senders", str(exc))
            return

        self.status_var.set("Loading unread emails from the last 2 days...")

        def _load() -> None:
            try:
                items = self.client.fetch_messages(self.allowed_senders)
                self.root.after(0, lambda: self._populate_messages(items))
            except Exception as exc:  # noqa: BLE001
                self.root.after(0, lambda: messagebox.showerror("Load error", str(exc)))
                self.root.after(0, lambda: self.status_var.set("Failed to load emails."))

        threading.Thread(target=_load, daemon=True).start()

    def _populate_messages(self, items: list[EmailMessage]) -> None:
        self.messages = items
        self.tree.delete(*self.tree.get_children())

        for index, item in enumerate(items):
            self.tree.insert(
                "",
                tk.END,
                iid=str(index),
                values=(item.sender, item.subject, item.received.replace("T", " ").replace("Z", "")),
            )

        self.preview.delete("1.0", tk.END)
        self.reply_subject.set("")
        self.reply_body.delete("1.0", tk.END)
        self.status_var.set(f"Loaded {len(items)} unread emails from allowed senders (last 2 days).")

    def on_select_message(self, _event: object) -> None:
        selection = self.tree.selection()
        if not selection:
            return

        selected = self.messages[int(selection[0])]
        self.preview.delete("1.0", tk.END)
        self.preview.insert(
            tk.END,
            f"From: {selected.sender}\nSubject: {selected.subject}\nReceived: {selected.received}\n\n{selected.body_preview}",
        )
        self.reply_subject.set(f"Re: {selected.subject}")

    def send_reply(self) -> None:
        if not self.client:
            messagebox.showerror("Not signed in", "Sign in before sending email.")
            return

        selection = self.tree.selection()
        if not selection:
            messagebox.showerror("No email selected", "Select an email to reply to.")
            return

        selected = self.messages[int(selection[0])]
        recipient = selected.sender.lower()
        if recipient not in self.allowed_senders:
            messagebox.showerror("Recipient blocked", "You can only send replies to the allowed senders list.")
            return

        subject = self.reply_subject.get().strip()
        body = self.reply_body.get("1.0", tk.END).strip()
        if not subject or not body:
            messagebox.showerror("Missing content", "Reply subject and body are required.")
            return

        self.status_var.set("Sending reply...")

        def _send() -> None:
            try:
                self.client.send_email(recipient, subject, body)
                self.root.after(0, lambda: messagebox.showinfo("Sent", "Reply sent successfully."))
                self.root.after(0, lambda: self.status_var.set("Reply sent."))
            except Exception as exc:  # noqa: BLE001
                self.root.after(0, lambda: messagebox.showerror("Send error", str(exc)))
                self.root.after(0, lambda: self.status_var.set("Failed to send reply."))

        threading.Thread(target=_send, daemon=True).start()


if __name__ == "__main__":
    app_root = tk.Tk()
    FocusMailApp(app_root)
    app_root.mainloop()
