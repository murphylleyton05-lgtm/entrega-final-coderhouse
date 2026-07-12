"""Subida a YouTube con la Data API v3 (OAuth). Importa perezoso y opcional.

Requisitos (una sola vez):
  1. En Google Cloud Console: crear proyecto, habilitar "YouTube Data API v3".
  2. Crear credenciales OAuth 2.0 tipo "Desktop app" y descargar el JSON.
  3. Guardarlo como credentials/client_secret.json en este proyecto.
La primera subida abre el navegador para autorizar y guarda credentials/token.json.
"""
from __future__ import annotations

import json
from pathlib import Path

SCOPES = ["https://www.googleapis.com/auth/youtube.upload"]


def _service(cred_dir: Path):
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build

    client_secret = cred_dir / "client_secret.json"
    token_file = cred_dir / "token.json"
    if not client_secret.exists():
        raise FileNotFoundError(
            f"Falta {client_secret}. Descargá las credenciales OAuth de Google Cloud "
            "(YouTube Data API v3, tipo Desktop app) y guardalas ahí."
        )

    creds = None
    if token_file.exists():
        creds = Credentials.from_authorized_user_file(str(token_file), SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(client_secret), SCOPES)
            creds = flow.run_local_server(port=0)
        token_file.write_text(creds.to_json(), encoding="utf-8")
    return build("youtube", "v3", credentials=creds)


def upload_video(cfg: dict, video_path: str, meta: dict) -> str:
    """Sube el video y devuelve el ID/URL. Lanza excepción si faltan dependencias."""
    try:
        from googleapiclient.http import MediaFileUpload
    except Exception as exc:  # pragma: no cover
        raise RuntimeError(
            "Faltan dependencias de Google. Instalá: "
            "pip install google-api-python-client google-auth-oauthlib"
        ) from exc

    cred_dir = Path(cfg["_root"]) / "credentials"
    cred_dir.mkdir(exist_ok=True)
    youtube = _service(cred_dir)

    body = {
        "snippet": {
            "title": meta["titulo"],
            "description": meta["descripcion"],
            "tags": meta.get("tags", []),
            "categoryId": str(cfg["youtube"].get("categoria_id", "27")),
        },
        "status": {
            "privacyStatus": cfg["youtube"].get("privacidad", "private"),
            "selfDeclaredMadeForKids": False,
        },
    }
    media = MediaFileUpload(video_path, chunksize=-1, resumable=True, mimetype="video/mp4")
    request = youtube.videos().insert(part="snippet,status", body=body, media_body=media)

    response = None
    while response is None:
        status, response = request.next_chunk()
        if status:
            print(f"  [YT] Subiendo... {int(status.progress() * 100)}%")
    vid = response["id"]
    url = f"https://youtu.be/{vid}"
    print(f"  [YT] Publicado ({cfg['youtube'].get('privacidad')}): {url}")
    return url
