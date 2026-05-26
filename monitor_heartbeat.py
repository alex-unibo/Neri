"""
Poller esterno dell'heartbeat del server centrale NeriCatering.

Progettato per girare su GitHub Actions (cron ogni 5 minuti) ma puo' essere
eseguito anche localmente.

Logica:
1. Legge sistema/server_heartbeat su Firestore.
2. Calcola eta' di last_seen. Se status != "running" oppure eta' > soglia
   (default 10 min), il server e' considerato offline.
3. Confronta con sistema/heartbeat_monitor_state.last_known_status:
   - Stato cambiato online -> offline: invia alert Telegram "OFFLINE".
   - Stato cambiato offline -> online: invia alert Telegram "RECUPERATO".
   - Stato invariato: nessuna notifica (no spam).
4. Aggiorna sistema/heartbeat_monitor_state.

Variabili d'ambiente richieste:
- FIREBASE_CREDENTIALS_JSON  Contenuto JSON delle credenziali Firebase
- TELEGRAM_BOT_TOKEN          Token del bot Telegram
- TELEGRAM_CHAT_ID            Chat destinatario degli alert
- ALERT_AFTER_MINUTES         (opzionale, default 10) eta' max del heartbeat

NB sulla timezone: il server scrive last_seen come datetime locale naive
(Europe/Rome). Il poller normalizza in UTC per il confronto, indipendente
dalla timezone del runner GitHub Actions (UTC).
"""

import os
import sys
import json
import tempfile
import traceback
import urllib.request
import urllib.parse
from datetime import datetime, timezone

try:
    from zoneinfo import ZoneInfo
except ImportError:
    # Python 3.8 fallback
    from backports.zoneinfo import ZoneInfo

import firebase_admin
from firebase_admin import credentials, firestore


SERVER_TZ = ZoneInfo("Europe/Rome")
UTC = ZoneInfo("UTC")

HEARTBEAT_DOC = ("sistema", "server_heartbeat")
STATE_DOC = ("sistema", "heartbeat_monitor_state")


def _env(name, default=None, required=True):
    val = os.environ.get(name, default)
    if required and not val:
        print(f"[FATAL] Variabile d'ambiente mancante: {name}", file=sys.stderr)
        sys.exit(1)
    return val


def send_telegram(token, chat_id, text):
    """Invia messaggio Telegram. Solleva eccezione se fallisce."""
    url = f"https://api.telegram.org/bot{token}/sendMessage"
    data = urllib.parse.urlencode({
        "chat_id": chat_id,
        "text": text,
        "parse_mode": "HTML",
        "disable_web_page_preview": "true",
    }).encode("utf-8")
    req = urllib.request.Request(url, data=data, method="POST")
    with urllib.request.urlopen(req, timeout=15) as resp:
        if resp.status != 200:
            raise RuntimeError(f"Telegram HTTP {resp.status}")
        return True


def init_firestore(creds_json):
    """Inizializza Firestore client da credenziali in stringa JSON."""
    with tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False, encoding="utf-8") as f:
        f.write(creds_json)
        temp_path = f.name
    try:
        cred = credentials.Certificate(temp_path)
        firebase_admin.initialize_app(cred)
        return firestore.client()
    finally:
        try:
            os.unlink(temp_path)
        except Exception:
            pass


def parse_server_timestamp(ts_str):
    """Converte un ISO datetime naive (Europe/Rome) in tz-aware UTC."""
    naive = datetime.fromisoformat(ts_str)
    if naive.tzinfo is None:
        naive = naive.replace(tzinfo=SERVER_TZ)
    return naive.astimezone(UTC)


def main():
    bot_token = _env("TELEGRAM_BOT_TOKEN")
    chat_id = _env("TELEGRAM_CHAT_ID")
    creds_json = _env("FIREBASE_CREDENTIALS_JSON")
    alert_after_minutes = int(_env("ALERT_AFTER_MINUTES", default="10", required=False))

    db = init_firestore(creds_json)
    now_utc = datetime.now(tz=UTC)

    # === 1. Leggi heartbeat ===
    hb_ref = db.collection(HEARTBEAT_DOC[0]).document(HEARTBEAT_DOC[1])
    hb_snap = hb_ref.get()

    if not hb_snap.exists:
        print("[WARN] Documento sistema/server_heartbeat non esiste su Firestore.")
        print("       Il server non e' mai stato avviato con la versione che scrive il heartbeat.")
        return 0

    hb = hb_snap.to_dict() or {}
    last_seen_str = hb.get("last_seen")
    server_status_field = hb.get("status", "unknown")
    hostname = hb.get("hostname", "?")

    if not last_seen_str:
        print("[WARN] Heartbeat senza campo last_seen. Salto.")
        return 0

    try:
        last_seen = parse_server_timestamp(last_seen_str)
    except Exception as e:
        print(f"[WARN] Impossibile parsare last_seen='{last_seen_str}': {e}")
        return 0

    age_seconds = (now_utc - last_seen).total_seconds()
    age_minutes = age_seconds / 60.0

    is_stale = age_seconds > alert_after_minutes * 60
    is_status_stopped = server_status_field != "running"

    is_online_now = (not is_stale) and (not is_status_stopped)
    new_status = "online" if is_online_now else "offline"

    # === 2. Leggi stato del poller ===
    state_ref = db.collection(STATE_DOC[0]).document(STATE_DOC[1])
    state_snap = state_ref.get()
    state = (state_snap.to_dict() or {}) if state_snap.exists else {}
    last_known_status = state.get("last_known_status", "online")

    print(f"Heartbeat last_seen      : {last_seen_str} (eta': {age_minutes:.1f} min)")
    print(f"Heartbeat status field   : {server_status_field}")
    print(f"Stato calcolato adesso   : {new_status}")
    print(f"Ultimo stato notificato  : {last_known_status}")
    print(f"Soglia alert             : {alert_after_minutes} min")

    # === 3. Edge-triggered alert ===
    if new_status != last_known_status:
        if new_status == "offline":
            text = (
                f"\U0001F534 <b>SERVER OFFLINE</b>\n"
                f"\U0001F5A5️ {hostname}\n"
                f"\U0001F552 Ultimo segno di vita: <b>{age_minutes:.0f} min fa</b>\n"
                f"\U0001F4CB Ultimo stato dichiarato: {server_status_field}"
            )
        else:
            text = (
                f"\U0001F7E2 <b>SERVER RECUPERATO</b>\n"
                f"\U0001F5A5️ {hostname}\n"
                f"\U0001F552 Heartbeat fresco: {age_minutes:.1f} min fa\n"
                f"\U0001F4CB Stato: {server_status_field}"
            )

        try:
            send_telegram(bot_token, chat_id, text)
            print(f"[OK] Alert Telegram inviato: {new_status.upper()}")
        except Exception as e:
            print(f"[ERR] Invio Telegram fallito: {e}")
            print("      Stato NON aggiornato. Si riprovera' al prossimo cron.")
            return 1

        state_ref.set({
            "last_known_status": new_status,
            "last_change_utc": now_utc.isoformat(),
            "last_check_utc": now_utc.isoformat(),
            "alert_after_minutes": alert_after_minutes,
        }, merge=True)
    else:
        state_ref.set({
            "last_check_utc": now_utc.isoformat(),
        }, merge=True)
        print(f"[OK] Stato invariato ({new_status}), nessuna notifica.")

    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as e:
        print(f"[FATAL] {type(e).__name__}: {e}", file=sys.stderr)
        traceback.print_exc()
        sys.exit(1)
