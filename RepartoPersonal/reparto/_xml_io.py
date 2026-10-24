"""Utilidades compartidas de lectura/escritura XML segura y atómica para archivos OpenXML (.xlsx).

Garantiza:
1. Preservación estricta de prefijos de espacios de nombres (xmlns:mc, xmlns:x14ac, etc.)
   para evitar avisos de corrupción o reparación al abrir en Microsoft Excel.
2. Escritura atómica mediante archivo temporal y reemplazo seguro.
3. Copias de seguridad rotativas con sello temporal (además de .bak inmediato).
4. Mensajes claros ante bloqueo de archivo (PermissionError en Windows cuando Excel lo tiene abierto).
"""

from datetime import datetime
from pathlib import Path
import re
import shutil
import xml.etree.ElementTree as ET
import zipfile

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

KNOWN_NAMESPACES = {
    "": MAIN_NS,
    "r": REL_NS,
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "x14ac": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
    "xr": "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
    "xr2": "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2",
    "xr3": "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3",
}


def register_xml_namespaces(xml_bytes: bytes) -> dict[str, str]:
    """Registra en ElementTree todos los espacios de nombres declarados en el XML original."""
    for prefix, uri in KNOWN_NAMESPACES.items():
        ET.register_namespace(prefix, uri)

    found: dict[str, str] = dict(KNOWN_NAMESPACES)
    text_head = xml_bytes[:4096].decode("utf-8", errors="ignore")
    for match in re.finditer(r'xmlns(?::([A-Za-z0-9_.-]+))?\s*=\s*["\']([^"\']+)["\']', text_head):
        prefix = match.group(1) or ""
        uri = match.group(2)
        found[prefix] = uri
        ET.register_namespace(prefix, uri)
    return found


def parse_xml_preserving_namespaces(xml_bytes: bytes) -> tuple[ET.Element, dict[str, str]]:
    """Parsea bytes XML registrando previamente todos sus prefijos de espacio de nombres."""
    ns_map = register_xml_namespaces(xml_bytes)
    root = ET.fromstring(xml_bytes)
    return root, ns_map


def serialize_xml_preserving_namespaces(
    root: ET.Element,
    ns_map: dict[str, str] | None = None,
) -> bytes:
    """Serializa el árbol XML asegurando que todos los prefijos de mc:Ignorable estén declarados."""
    if ns_map:
        for prefix, uri in ns_map.items():
            ET.register_namespace(prefix, uri)
    else:
        for prefix, uri in KNOWN_NAMESPACES.items():
            ET.register_namespace(prefix, uri)

    raw_bytes = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    xml_text = raw_bytes.decode("utf-8")

    # Verificar si el nodo raíz contiene mc:Ignorable="..." con prefijos que ET haya omitido
    # por no tener nodos hijos activos con ese prefijo.
    ignorable_match = re.search(r'\b(?:\w+:)?Ignorable\s*=\s*["\']([^"\']+)["\']', xml_text[:2048])
    if ignorable_match:
        prefixes = ignorable_match.group(1).split()
        active_ns = ns_map or KNOWN_NAMESPACES
        missing_decls = []
        for p in prefixes:
            if f'xmlns:{p}=' not in xml_text[:2048] and p in active_ns:
                missing_decls.append(f'xmlns:{p}="{active_ns[p]}"')
        if missing_decls:
            # Inyectar las declaraciones faltantes justo antes del cierre de la etiqueta raíz
            decl_str = " " + " ".join(missing_decls)
            xml_text = re.sub(
                r'(<worksheet\b[^>]*)(>)',
                rf'\1{decl_str}\2',
                xml_text,
                count=1,
            )

    return xml_text.encode("utf-8")


def create_rotating_backup(path: Path, max_backups: int = 10) -> Path:
    """Crea copia .bak inmediata y copia rotativa con fecha en subcarpeta 'copias/'."""
    path = Path(path)
    bak_path = path.with_suffix(path.suffix + ".bak")
    shutil.copy2(path, bak_path)

    try:
        backup_dir = path.parent / "copias"
        backup_dir.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        rot_path = backup_dir / f"{path.stem}_{stamp}{path.suffix}"
        shutil.copy2(path, rot_path)

        # Conservar solo las últimas max_backups copias de este archivo
        pattern = f"{path.stem}_*{path.suffix}"
        existing = sorted(backup_dir.glob(pattern), key=lambda p: p.stat().st_mtime)
        while len(existing) > max_backups:
            oldest = existing.pop(0)
            oldest.unlink(missing_ok=True)
    except OSError:
        # Si el directorio es de solo lectura para subcarpetas, el .bak principal ya quedó creado
        pass

    return bak_path


def atomic_write_zip(path: Path, files: dict[str, bytes]) -> None:
    """Escribe un diccionario de archivos en un contenedor ZIP (.xlsx) de forma atómica."""
    path = Path(path)
    temp_path = path.with_suffix(path.suffix + ".tmp")
    try:
        with zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for name, content in files.items():
                zout.writestr(name, content)
        temp_path.replace(path)
    except PermissionError as exc:
        temp_path.unlink(missing_ok=True)
        raise PermissionError(
            f"No se puede guardar '{path.name}' porque está abierto en otro programa "
            f"(por ejemplo, Microsoft Excel). Cierre el archivo e inténtelo de nuevo."
        ) from exc
    except Exception:
        temp_path.unlink(missing_ok=True)
        raise
