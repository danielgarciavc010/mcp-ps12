"""
MCP Server - ivanti-itsm
Transporte: stdio

Arranque:
    python server.py
    -> servidor MCP por stdin/stdout

Herramientas:
    - list_business_objects
    - get_business_object
    - create_business_object
    - create_incident_simple
    - update_business_object
    - delete_business_object
    - run_quick_action
    - list_fields
    - run_saved_search
    - get_related_objects
    - link_business_objects
    - get_metadata
    - find_user_info
    - list_service_catalog_templates
    - get_template_fields
    - get_field_valid_values
    - create_service_request
    - upload_attachment

Referencia de la API:
https://help.ivanti.com/ht/help/en_US/ISM/2021/admin/Content/Configure/API/RestAPI-Introduction.htm
"""

import os
import re
import sys
import unicodedata
import pathlib as _pathlib
from typing import Any, Optional

import httpx
from dotenv import load_dotenv
from fastmcp import FastMCP

try:
    import truststore
    truststore.inject_into_ssl()
except ImportError:
    pass

# Config
load_dotenv()

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stdin, "reconfigure"):
    sys.stdin.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

TENANT_URL = os.getenv("IVANTI_TENANT_URL", "").rstrip("/")
API_KEY    = os.getenv("IVANTI_API_KEY", "")
TIMEOUT    = 30.0

if not TENANT_URL or not API_KEY:
    raise RuntimeError(
        "Configura IVANTI_TENANT_URL e IVANTI_API_KEY en el archivo .env "
        "(copia .env.example como .env y rellena la API Key)."
    )

HEADERS = {
    "Authorization": f"rest_api_key={API_KEY}",
    "Content-Type":  "application/json",
}

mcp = FastMCP(
    name="ivanti-itsm",
    instructions=(
        "Herramientas para gestionar Ivanti Neurons ITSM: consultar, crear, "
        "actualizar y eliminar business objects (incidentes, solicitudes de "
        "servicio, problemas, cambios, empleados), ejecutar quick actions, "
        "gestionar solicitudes del catalogo de servicios y subir adjuntos."
    ),
)


def _clean_unicode(obj: Any) -> Any:
    if isinstance(obj, str):
        s = obj.replace(chr(160), " ").replace("\u00a0", " ")
        s = unicodedata.normalize("NFKD", s)
        s = "".join(c for c in s if not unicodedata.combining(c))
        return s.encode("ascii", errors="replace").decode("ascii")
    elif isinstance(obj, dict):
        return {_clean_unicode(k): _clean_unicode(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [_clean_unicode(elem) for elem in obj]
    elif isinstance(obj, tuple):
        return tuple(_clean_unicode(elem) for elem in obj)
    return obj


def _error(code: str, message: str) -> dict:
    return {"ok": False, "error": {"code": code, "message": message}}


def _odata_literal(value: Any) -> str:
    if value is None:
        return "null"
    if isinstance(value, bool):
        return str(value).lower()
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return str(value)
    escaped = str(value).replace("'", "''")
    return f"'{escaped}'"


def _build_odata_filter(filters: list[dict[str, Any]]) -> str | None:
    allowed_operators = {
        "eq", "ne", "gt", "ge", "lt", "le",
        "contains", "startswith", "endswith",
    }
    expressions = []
    for item in filters:
        field = str(item.get("field", "")).strip()
        operator = str(item.get("operator", "")).lower().strip()
        if (
            not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", field)
            or operator not in allowed_operators
            or "value" not in item
        ):
            return None
        value = _odata_literal(item["value"])
        if operator in {"contains", "startswith", "endswith"}:
            expressions.append(f"{operator}({field}, {value})")
        else:
            expressions.append(f"{field} {operator} {value}")
    return " and ".join(f"({expression})" for expression in expressions)


async def _request(
    method: str,
    endpoint: str,
    *,
    headers: dict | None = None,
    **kwargs,
) -> tuple[Any, dict | None]:
    req_headers = headers if headers is not None else HEADERS
    try:
        async with httpx.AsyncClient(
            base_url=TENANT_URL,
            headers=req_headers,
            timeout=TIMEOUT,
        ) as client:
            response = await client.request(method, endpoint, **kwargs)
    except httpx.TimeoutException:
        return None, _error("TIMEOUT", "El servicio no respondio a tiempo.")
    except httpx.RequestError as exc:
        return None, _error(
            "CONNECTION_ERROR", f"No se pudo conectar al servicio: {exc}."
        )

    if response.status_code >= 400:
        return None, _error(
            "HTTP_ERROR",
            f"Ivanti API error {response.status_code}: {response.text}",
        )

    try:
        return _clean_unicode(response.json()), None
    except Exception:
        return _clean_unicode(response.text), None


# Tools - Business Objects (CRUD)

@mcp.tool()
async def list_business_objects(
    object_name: str,
    top: int = 10,
    select: Optional[str] = None,
    search: Optional[str] = None,
    filters: Optional[list[dict[str, Any]]] = None,
) -> dict:
    """
    Lista registros de un business object de Ivanti (p.ej. 'incidents',
    'employees', 'serviceReqs', 'problems', 'changes'). El nombre del
    objeto siempre va en plural.

    Args:
        object_name: nombre del business object en plural (ej: 'incidents').
        top: numero maximo de registros a devolver (max. 25).
        select: campos a devolver separados por coma (ej: "Subject,Status").
        search: palabra clave de busqueda de texto libre ($search).
        filters: filtros estructurados. Cada elemento debe incluir
            "field", "operator" y "value". Operadores: eq, ne, gt, ge,
            lt, le, contains, startswith, endswith. Se combinan con AND.

    Returns:
        {"ok": true, "data": { ...registros del servicio... }}
    """
    DEFAULT_SELECT = (
        "RecId,IncidentNumber,Subject,Status,Priority,"
        "Category,CreatedBy,CreatedDateTime"
    )
    MAX_TOP = 25

    if filters:
        generated_filter = _build_odata_filter(filters)
        if generated_filter is None:
            return _error(
                "INVALID_FILTER",
                "Cada filtro requiere field, operator valido y value.",
            )
    else:
        generated_filter = None

    params: dict[str, Any] = {"$top": min(top, MAX_TOP)}
    if generated_filter:
        params["$filter"] = generated_filter
    if select:
        params["$select"] = select
    elif object_name.lower() in ("incidents", "incident"):
        params["$select"] = DEFAULT_SELECT
    if search:
        params["$search"] = search

    raw, err = await _request(
        "GET", f"/api/odata/businessobject/{object_name}", params=params
    )
    if err:
        return err

    return {"ok": True, "data": raw}


@mcp.tool()
async def get_business_object(object_name: str, rec_id: str) -> dict:
    """
    Obtiene un registro concreto de un business object por su RecId.

    Args:
        object_name: nombre del business object en plural (ej: 'incidents').
        rec_id: RecId (identificador unico) del registro.

    Returns:
        {"ok": true, "data": { ...campos del registro... }}
    """
    raw, err = await _request(
        "GET", f"/api/odata/businessobject/{object_name}('{rec_id}')"
    )
    if err:
        return err

    return {"ok": True, "data": raw}


@mcp.tool()
async def create_business_object(
    object_name: str, fields: dict[str, Any]
) -> dict:
    """
    Crea un nuevo registro de un business object.

    Args:
        object_name: nombre del business object en plural (ej: 'incidents').
        fields: diccionario con los campos y valores a establecer, ej:
            {"Subject": "No imprime", "Status": "Open", "Priority": 3}

    Returns:
        {"ok": true, "data": { ...registro creado... }}
    """
    raw, err = await _request(
        "POST", f"/api/odata/businessobject/{object_name}", json=fields
    )
    if err:
        return err

    return {"ok": True, "data": raw}


@mcp.tool()
async def create_incident_simple(
    subject: str,
    description: str,
    customer_name: str,
    urgency: str = "Medium",
    impact: str = "Medium",
    category: str = "Service Desk",
    service: str = "Service Desk",
) -> dict:
    """
    Crea un ticket de incidencia de forma segura, resolviendo las
    validaciones tecnicas de Ivanti. USALO siempre que el usuario pida
    crear un incidente o reportar un problema.

    Args:
        subject: Resumen breve del problema.
        description: Detalles completos del sintoma o problema.
        customer_name: Nombre completo, login o email del usuario afectado.
        urgency: Urgencia: 'High', 'Medium', 'Low'.
        impact: Impacto: 'High', 'Medium', 'Low'.
        category: Categoria del incidente.
        service: Servicio asociado.

    Returns:
        {"ok": true, "data": {"IncidentNumber": "...", "RecId": "...", ...}}
    """
    safe_customer = customer_name.replace("'", "''")
    search_filter = (
        f"contains(DisplayName, '{safe_customer}') or "
        f"contains(PrimaryEmail, '{safe_customer}') or "
        f"contains(LoginID, '{safe_customer}')"
    )

    raw_emp, err = await _request(
        "GET",
        "/api/odata/businessobject/employees",
        params={
            "$filter": search_filter,
            "$top": 1,
            "$select": "RecId,DisplayName,PrimaryEmail,LoginID,Team,Team_Valid",
        },
    )
    if err:
        return _error(
            "EMPLOYEE_SEARCH_ERROR",
            f"Error buscando empleado '{customer_name}': {err['error']['message']}",
        )

    values = raw_emp.get("value", []) if isinstance(raw_emp, dict) else []
    if not values:
        return _error(
            "EMPLOYEE_NOT_FOUND",
            f"No se encontro un empleado coincidente con '{customer_name}'. "
            "Pide al usuario que verifique el nombre, login o email.",
        )

    employee = values[0]
    customer_rec_id = employee["RecId"]
    customer_display = employee.get("DisplayName", customer_name)
    employee_team = employee.get("Team", "Service Desk")

    fields = {
        "Subject": subject,
        "Symptom": description,
        "Urgency": urgency,
        "Impact": impact,
        "Status": "Logged",
        "Category": category,
        "Service": service,
        "ProfileLink_RecID": customer_rec_id,
        "ProfileLink_Category": "Employee",
        "Customer": customer_display,
        "OwnerTeam": "Service Desk",
        "Owner": "Admin",
        "Team": employee_team if employee_team else "Service Desk",
        "Source": "Phone",
    }

    raw, err = await _request(
        "POST", "/api/odata/businessobject/incidents", json=fields
    )
    if err:
        return err

    inc_number = raw.get("IncidentNumber", "desconocido") if isinstance(raw, dict) else "desconocido"
    return {
        "ok": True,
        "data": {
            "message": f"Incidente #{inc_number} creado correctamente para {customer_display}.",
            "IncidentNumber": inc_number,
            "RecId": raw.get("RecId", "") if isinstance(raw, dict) else "",
            "Status": raw.get("Status", "Logged") if isinstance(raw, dict) else "Logged",
        },
    }


@mcp.tool()
async def update_business_object(
    object_name: str, rec_id: str, fields: dict[str, Any]
) -> dict:
    """
    Actualiza campos de un registro existente.

    Args:
        object_name: nombre del business object en plural (ej: 'incidents').
        rec_id: RecId del registro a actualizar.
        fields: campos a modificar, ej: {"Status": "Closed"}

    Returns:
        {"ok": true, "data": { ...registro actualizado... }}
    """
    raw, err = await _request(
        "PATCH",
        f"/api/odata/businessobject/{object_name}('{rec_id}')",
        json=fields,
    )
    if err:
        return err

    return {"ok": True, "data": raw if raw else {"status": "ok"}}


@mcp.tool()
async def delete_business_object(object_name: str, rec_id: str) -> dict:
    """
    Elimina un registro de un business object.

    Args:
        object_name: nombre del business object en plural (ej: 'incidents').
        rec_id: RecId del registro a eliminar.

    Returns:
        {"ok": true, "data": {"status": "deleted", "rec_id": "..."}}
    """
    _, err = await _request(
        "DELETE", f"/api/odata/businessobject/{object_name}('{rec_id}')"
    )
    if err:
        return err

    return {"ok": True, "data": {"status": "deleted", "rec_id": rec_id}}


# Tools - Quick Actions & Saved Searches

@mcp.tool()
async def run_quick_action(
    object_name: str, rec_id: str, quick_action_name: str
) -> dict:
    """
    Ejecuta una quick action sobre un registro (ej: cerrar, resolver o
    asignar un incidente).

    Args:
        object_name: nombre del business object en plural (ej: 'incidents').
        rec_id: RecId del registro sobre el que aplicar la accion.
        quick_action_name: nombre de la quick action (ej: 'Close_Incident').

    Returns:
        {"ok": true, "data": { ...resultado de la accion... }}
    """
    raw, err = await _request(
        "POST",
        f"/api/odata/businessobject/{object_name}('{rec_id}')/{quick_action_name}",
    )
    if err:
        return err

    return {"ok": True, "data": raw if raw else {"status": "ok"}}


@mcp.tool()
async def run_saved_search(
    object_name: str,
    saved_search_name: str,
    action_id: str = "",
) -> dict:
    """
    Ejecuta una Saved Search ya configurada en Ivanti.

    Args:
        object_name: nombre del business object en plural (ej: 'incidents').
        saved_search_name: nombre exacto de la saved search.
        action_id: identificador opcional de la accion asociada.

    Returns:
        {"ok": true, "data": { ...resultados de la busqueda... }}
    """
    raw, err = await _request(
        "GET",
        f"/api/odata/businessobject/{object_name}/{saved_search_name}",
        params={"ActionId": action_id},
    )
    if err:
        return err

    return {"ok": True, "data": raw}


# Tools - Metadata & Discovery

@mcp.tool()
async def get_metadata(object_name: str) -> dict:
    """
    Devuelve los metadatos (campos, tipos, relaciones...) de un business
    object como texto XML plano.

    Args:
        object_name: nombre del business object en plural (ej: 'incidents').

    Returns:
        {"ok": true, "data": "<xml>...metadatos...</xml>"}
    """
    async with httpx.AsyncClient(
        base_url=TENANT_URL, headers=HEADERS, timeout=TIMEOUT
    ) as client:
        try:
            resp = await client.get(f"/api/odata/{object_name}/$metadata")
        except httpx.TimeoutException:
            return _error("TIMEOUT", "El servicio no respondio a tiempo.")
        except httpx.RequestError as exc:
            return _error("CONNECTION_ERROR", f"No se pudo conectar: {exc}.")

    if resp.status_code >= 400:
        return _error("HTTP_ERROR", f"Ivanti API error {resp.status_code}: {resp.text}")

    raw_bytes = resp.content
    if raw_bytes[:2] in (b"\xff\xfe", b"\xfe\xff"):
        text = raw_bytes.decode("utf-16")
    else:
        text = resp.text

    return _clean_unicode({"ok": True, "data": text})


@mcp.tool()
async def list_fields(object_name: str) -> dict:
    """
    Devuelve la lista de campos (nombre y tipo) de un business object,
    parseada desde el XML de metadatos. Mas compacta que get_metadata().

    Args:
        object_name: nombre del business object en plural (ej: 'incidents').

    Returns:
        {"ok": true, "data": {"object_name": "...", "fields": [...]}}
    """
    result = await get_metadata(object_name)
    if not result.get("ok"):
        return result

    xml_text = result["data"]

    candidates = [object_name, object_name.rstrip("s")]
    entity_block = None
    matched_name = None

    for candidate in candidates:
        pattern = rf'<EntityType Name="{re.escape(candidate)}">'
        match = re.search(pattern, xml_text, flags=re.IGNORECASE)
        if match:
            start = match.start()
            end_pos = xml_text.find("</EntityType>", start)
            if end_pos == -1:
                continue
            end = end_pos + len("</EntityType>")
            entity_block = xml_text[start:end]
            matched_name = candidate
            break

    if entity_block is None:
        available = sorted(
            set(re.findall(r'<EntityType Name="([^"]+)"', xml_text))
        )
        return _error(
            "ENTITY_NOT_FOUND",
            f"No se encontro EntityType para '{object_name}'. "
            f"Tipos disponibles: {available[:50]}",
        )

    props = re.findall(
        r'<Property Name="([^"]+)" Type="([^"]+)"', entity_block
    )

    return {
        "ok": True,
        "data": {
            "object_name": object_name,
            "entity_type_matched": matched_name,
            "field_count": len(props),
            "fields": [{"name": n, "type": t} for n, t in props],
        },
    }


# Tools - Relationships

@mcp.tool()
async def get_related_objects(
    object_name: str, rec_id: str, relationship_name: str
) -> dict:
    """
    Devuelve los objetos relacionados con un registro.

    Args:
        object_name: nombre del business object en plural (ej: 'incidents').
        rec_id: RecId del registro principal.
        relationship_name: nombre de la relacion (ej: 'IncidentContainsJournal').

    Returns:
        {"ok": true, "data": { ...objetos relacionados... }}
    """
    raw, err = await _request(
        "GET",
        f"/api/odata/businessobject/{object_name}('{rec_id}')/{relationship_name}",
    )
    if err:
        return err

    return {"ok": True, "data": raw}


@mcp.tool()
async def link_business_objects(
    object_name: str,
    rec_id: str,
    relationship_name: str,
    related_rec_id: str,
) -> dict:
    """
    Vincula dos registros mediante una relacion.

    Args:
        object_name: nombre del business object principal en plural.
        rec_id: RecId del registro principal.
        relationship_name: nombre de la relacion.
        related_rec_id: RecId del registro relacionado a vincular.

    Returns:
        {"ok": true, "data": {"status": "linked", ...}}
    """
    _, err = await _request(
        "PATCH",
        f"/api/odata/businessobject/{object_name}('{rec_id}')/"
        f"{relationship_name}('{related_rec_id}')/$Ref",
    )
    if err:
        return err

    return {
        "ok": True,
        "data": {
            "status": "linked",
            "rec_id": rec_id,
            "related_rec_id": related_rec_id,
        },
    }


# Tools - Service Catalog

@mcp.tool()
async def find_user_info(network_username: str) -> dict:
    """
    Busca los datos de un usuario/empleado necesarios para crear una
    Solicitud de Servicio: su RecId y su ubicacion.

    Args:
        network_username: nombre de usuario de red (login) del empleado.

    Returns:
        {"ok": true, "data": { ...datos del usuario... }}
    """
    raw, err = await _request(
        "GET",
        "/api/odata/businessobject/Frs_CompositeContract_Contacts",
        params={"$filter": f"NetworkUserName eq '{network_username}'"},
    )
    if err:
        return err

    return {"ok": True, "data": raw}


@mcp.tool()
async def list_service_catalog_templates(
    user_rec_id: str, name_contains: Optional[str] = None
) -> dict:
    """
    Lista las plantillas de Solicitud de Servicio publicadas en el catalogo
    disponibles para un usuario.

    Args:
        user_rec_id: RecId del usuario (obtenido con find_user_info).
        name_contains: texto opcional para filtrar plantillas por nombre.

    Returns:
        {"ok": true, "data": {"templates": [...]}}
    """
    raw, err = await _request(
        "GET", f"/api/rest/Template/{user_rec_id}/_All_"
    )
    if err:
        return err

    data = raw
    if name_contains and isinstance(data, list):
        needle = name_contains.lower()
        data = [
            t
            for t in data
            if needle in str(t.get("strName", "")).lower()
        ]

    return {"ok": True, "data": {"templates": data}}


@mcp.tool()
async def get_template_fields(
    subscription_id: str, customer_location: str
) -> dict:
    """
    Devuelve los campos/parametros de una plantilla de Solicitud de Servicio.

    Args:
        subscription_id: strSubscriptionId de la plantilla.
        customer_location: ubicacion del usuario (de find_user_info).

    Returns:
        {"ok": true, "data": { ...campos de la plantilla... }}
    """
    raw, err = await _request(
        "GET",
        f"/api/rest/ServiceRequest/PackageData/{subscription_id}/{customer_location}",
    )
    if err:
        return err

    return {"ok": True, "data": raw}


@mcp.tool()
async def get_field_valid_values(parameter_rec_id: str) -> dict:
    """
    Para un campo tipo desplegable/pick-list, devuelve los valores permitidos.

    Args:
        parameter_rec_id: RecID del parametro (de get_template_fields).

    Returns:
        {"ok": true, "data": { ...valores validos... }}
    """
    raw, err = await _request(
        "GET",
        f"/api/rest/ServiceRequest/{parameter_rec_id}/ValidationList",
    )
    if err:
        return err

    return {"ok": True, "data": raw}


@mcp.tool()
async def create_service_request(
    subscription_id: str,
    user_rec_id: str,
    customer_location: str,
    subject: str,
    symptom: str,
    parameters: Optional[list[dict[str, str]]] = None,
    local_offset_minutes: int = 0,
) -> dict:
    """
    Crea una Solicitud de Servicio real a partir de una plantilla del
    catalogo. Ejecuta esto SOLO despues de confirmar con el usuario.

    Args:
        subscription_id: strSubscriptionId de la plantilla.
        user_rec_id: RecId del solicitante (de find_user_info).
        customer_location: ubicacion del solicitante (de find_user_info).
        subject: asunto/titulo de la solicitud.
        symptom: descripcion/detalle de la solicitud.
        parameters: lista opcional de campos adicionales.
        local_offset_minutes: diferencia horaria en minutos respecto a UTC.

    Returns:
        {"ok": true, "data": { ...solicitud creada... }}
    """
    parameters = parameters or []
    parameters_payload: dict[str, str] = {}
    for p in parameters:
        rec_id = p["rec_id"]
        parameters_payload[f"par-{rec_id}"] = p["value"]
        if p.get("value_rec_id"):
            parameters_payload[f"par-{rec_id}-recId"] = p["value_rec_id"]

    payload = {
        "attachmentsToDelete": [],
        "attachmentsToUpload": [],
        "parameters": parameters_payload,
        "delayedFulfill": False,
        "formName": "ServiceReq.ResponsiveAnalyst.DefaultLayout",
        "saveReqState": False,
        "serviceReqData": {"Subject": subject, "Symptom": symptom},
        "strCustomerLocation": customer_location,
        "strUserId": user_rec_id,
        "subscriptionId": subscription_id,
        "localOffset": local_offset_minutes,
    }

    raw, err = await _request(
        "POST", "/api/rest/ServiceRequest/new", json=payload
    )
    if err:
        return err

    return {"ok": True, "data": raw}


# Tools - Adjuntos

@mcp.tool()
async def upload_attachment(
    object_name: str, rec_id: str, file_path: str
) -> dict:
    """
    Sube un archivo adjunto a un registro existente.

    Args:
        object_name: nombre del business object en SINGULAR
            (ej: 'incident', no 'incidents').
        rec_id: RecId del registro al que adjuntar el archivo.
        file_path: ruta local del archivo a subir.

    Returns:
        {"ok": true, "data": { ...respuesta del servicio... }}
    """
    path = _pathlib.Path(file_path)
    if not path.exists():
        return _error(
            "FILE_NOT_FOUND",
            f"No se encontro el archivo local: {file_path}",
        )

    upload_headers = {"Authorization": HEADERS["Authorization"]}

    try:
        async with httpx.AsyncClient(
            base_url=TENANT_URL, headers=upload_headers, timeout=60
        ) as client:
            with open(path, "rb") as f:
                response = await client.post(
                    "/api/rest/Attachment",
                    data={"ObjectID": rec_id, "ObjectType": object_name},
                    files={"File": (path.name, f)},
                )
    except httpx.TimeoutException:
        return _error("TIMEOUT", "El servicio no respondio a tiempo.")
    except httpx.RequestError as exc:
        return _error(
            "CONNECTION_ERROR", f"No se pudo conectar al servicio: {exc}."
        )

    if response.status_code >= 400:
        return _error(
            "HTTP_ERROR",
            f"Ivanti API error {response.status_code}: {response.text}",
        )

    try:
        return {"ok": True, "data": response.json()}
    except Exception:
        return {"ok": True, "data": response.text}


# Arranque
def main():
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
