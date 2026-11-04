"""
Módulo de Persistencia de Sesión de Archivos - Agente 007
Garantiza que los archivos Excel/CSV subidos por los usuarios en Telegram
sobrevivan a los reinicios, suspensiones ('spin-downs') y despliegues de Render.

Estrategia Híbrida de 3 Niveles:
1. Memoria RAM (Ultra rápido durante la ejecución continua).
2. Disco Persistente Local (/tmp/agente007_sessions/ o ./sessions/).
3. Supabase / Telegram CDN (Re-descarga mediante file_id permanente si se reinicia el servidor).
"""

import os
import json
import io
import time
import threading
import httpx
import pandas as pd

SESSIONS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sessions_cache")
os.makedirs(SESSIONS_DIR, exist_ok=True)

def _supabase_sync_thread(supabase_url: str, supabase_key: str, chat_id: int, file_id: str, file_name: str, active_sheet: str):
    """Sincroniza la sesión con Supabase en un hilo de fondo sin bloquear el hilo principal."""
    try:
        headers = {
            "apikey": supabase_key,
            "Authorization": f"Bearer {supabase_key}",
            "Content-Type": "application/json",
            "Prefer": "resolution=merge-duplicates"
        }
        payload = {
            "chat_id": str(chat_id),
            "file_id": file_id,
            "file_name": file_name,
            "active_sheet": active_sheet,
            "updated_at": "now()"
        }
        with httpx.Client(timeout=3.0) as client:
            client.post(f"{supabase_url}/rest/v1/sesiones_excel", headers=headers, json=payload)
    except Exception:
        pass

def save_user_session(chat_id: int, file_id: str, file_name: str, file_bytes: bytes, sheets: list, active_sheet: str, df: pd.DataFrame = None):
    """Guarda el archivo y sus metadatos en caché de disco y registra en Supabase si está disponible."""
    user_file_path = os.path.join(SESSIONS_DIR, f"{chat_id}_latest.xlsx")
    user_meta_path = os.path.join(SESSIONS_DIR, f"{chat_id}_meta.json")

    # 1. Guardar archivo binario en disco
    try:
        with open(user_file_path, "wb") as f:
            f.write(file_bytes)
    except Exception as e:
        print(f"⚠️ Error guardando sesión en disco: {e}", flush=True)

    # 2. Guardar metadatos en disco
    meta = {
        "chat_id": chat_id,
        "file_id": file_id,
        "file_name": file_name,
        "sheets": sheets,
        "active_sheet": active_sheet,
        "total_filas": len(df) if df is not None else 0,
        "timestamp": time.time()
    }
    try:
        with open(user_meta_path, "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"⚠️ Error guardando metadatos en disco: {e}", flush=True)

    # 3. Registro asíncrono en Supabase (hilo en segundo plano sin latencia)
    supabase_url = os.getenv("SUPABASE_URL", "").strip()
    supabase_key = os.getenv("SUPABASE_KEY", "").strip()
    if supabase_url and supabase_key:
        threading.Thread(
            target=_supabase_sync_thread,
            args=(supabase_url, supabase_key, chat_id, file_id, file_name, active_sheet),
            daemon=True
        ).start()

async def restore_user_session(chat_id: int, client: httpx.AsyncClient = None, bot_token: str = None) -> dict:
    """
    Restaura la sesión del usuario si la RAM fue borrada por un reinicio de Render.
    Retorna un diccionario con 'raw_bytes', 'filename', 'sheets', 'active_sheet', 'df' o None.
    """
    user_file_path = os.path.join(SESSIONS_DIR, f"{chat_id}_latest.xlsx")
    user_meta_path = os.path.join(SESSIONS_DIR, f"{chat_id}_meta.json")

    # Caso 1: Archivo y metadatos existen en disco
    if os.path.exists(user_file_path) and os.path.exists(user_meta_path):
        try:
            with open(user_file_path, "rb") as f:
                raw_bytes = f.read()
            with open(user_meta_path, "r", encoding="utf-8") as f:
                meta = json.load(f)

            if len(raw_bytes) > 0:
                return {
                    "raw_bytes": raw_bytes,
                    "filename": meta.get("file_name", "archivo.xlsx"),
                    "sheets": meta.get("sheets", ["Base_Datos"]),
                    "active_sheet": meta.get("active_sheet", "Base_Datos"),
                    "file_id": meta.get("file_id", "")
                }
        except Exception as e:
            print(f"⚠️ Error restaurando desde disco: {e}", flush=True)

    # Caso 2: Consultar Supabase para obtener file_id y re-descargar de Telegram
    try:
        supabase_url = os.getenv("SUPABASE_URL", "").strip()
        supabase_key = os.getenv("SUPABASE_KEY", "").strip()
        if supabase_url and supabase_key and client and bot_token:
            from app.db.supabase import get_supabase_client
            supabase = get_supabase_client()
            res = supabase.table("sesiones_excel").select("*").eq("chat_id", str(chat_id)).limit(1).execute()
            if res.data and len(res.data) > 0:
                record = res.data[0]
                fid = record.get("file_id")
                fname = record.get("file_name", "archivo.xlsx")
                if fid:
                    print(f"🔄 Re-descargando archivo desde CDN de Telegram usando file_id={fid}...", flush=True)
                    # Descargar de Telegram
                    t_resp = await client.get(f"https://api.telegram.org/bot{bot_token}/getFile?file_id={fid}")
                    if t_resp.status_code == 200:
                        fpath = t_resp.json().get("result", {}).get("file_path")
                        if fpath:
                            d_resp = await client.get(f"https://api.telegram.org/file/bot{bot_token}/{fpath}")
                            if d_resp.status_code == 200:
                                b = d_resp.content
                                # Guardar en disco local para acelerar próximas peticiones
                                save_user_session(chat_id, fid, fname, b, [record.get("active_sheet", "Base_Datos")], record.get("active_sheet", "Base_Datos"))
                                return {
                                    "raw_bytes": b,
                                    "filename": fname,
                                    "sheets": [record.get("active_sheet", "Base_Datos")],
                                    "active_sheet": record.get("active_sheet", "Base_Datos"),
                                    "file_id": fid
                                }
    except Exception as e_supa:
        pass

    return None
