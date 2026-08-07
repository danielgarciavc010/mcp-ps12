"""
Comprueba que la URL del tenant y la API Key de Ivanti Neurons for ITSM
funcionan ANTES de construir el servidor MCP.

Uso:
    python test_connection.py
"""
import os
import sys

import truststore

truststore.inject_into_ssl()  # usa el almacén de certificados de Windows/macOS/Linux
import httpx
from dotenv import load_dotenv

load_dotenv()

TENANT_URL = os.getenv("IVANTI_TENANT_URL", "").rstrip("/")
API_KEY = os.getenv("IVANTI_API_KEY", "")


def main() -> None:
    if not TENANT_URL or not API_KEY:
        print("❌ Faltan IVANTI_TENANT_URL o IVANTI_API_KEY en el archivo .env")
        sys.exit(1)

    url = f"{TENANT_URL}/api/odata/businessobject/incidents"
    headers = {
        "Authorization": f"rest_api_key={API_KEY}",
        "Content-Type": "application/json",
    }
    params = {"$top": 1}

    print(f"Tenant : {TENANT_URL}")
    print(f"URL    : {url}")
    print("Probando conexión...\n")

    try:
        resp = httpx.get(url, headers=headers, params=params, timeout=15)
    except httpx.RequestError as e:
        print(f"❌ Error de red al conectar con el tenant: {e}")
        sys.exit(1)

    print(f"Status code: {resp.status_code}")

    if resp.status_code == 200:
        print("✅ Conexión y autenticación correctas.\n")
        print(resp.text[:1000])
    elif resp.status_code in (401, 403):
        print("❌ La API Key fue rechazada (no autorizada).")
        print("   Revisa en Ivanti: Configure > Security Controls > API Keys")
        print("   -> que la key esté 'Activated' y tenga 'On Behalf Of' / 'In Role' configurados.")
        print(resp.text)
    elif resp.status_code == 404:
        print("❌ 404 - Revisa que IVANTI_TENANT_URL sea correcta y no acabe en '/'.")
        print(resp.text)
    else:
        print("⚠️ Respuesta inesperada:")
        print(resp.text[:1000])


if __name__ == "__main__":
    main()
