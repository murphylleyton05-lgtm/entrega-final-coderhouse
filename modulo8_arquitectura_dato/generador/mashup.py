"""
Empaqueta el codigo M dentro del .xlsx como parte DataMashup.

Un .xlsx guarda las consultas de Power Query en `customXml/item1.xml`, dentro
de un elemento `<DataMashup>` cuyo contenido es un binario en base64. El
formato de ese binario esta documentado por Microsoft en [MS-QDEFF] (Query
Definition File Format) y es simple:

    Version                     4 bytes, entero little-endian, vale 0
    Longitud de Package Parts   4 bytes
    Package Parts               un zip con el codigo M
    Longitud de Permissions     4 bytes
    Permissions                 XML de permisos
    Longitud de Metadata        4 bytes
    Metadata                    version + XML de metadatos + contenido
    Longitud de Bindings        4 bytes
    Permission Bindings         checksum; la spec permite un unico byte 0x00

El zip de Package Parts contiene tres archivos:

    [Content_Types].xml     tipos de contenido del paquete
    Config/Package.xml      version y cultura del motor de mashup
    Formulas/Section1.m     el codigo M (una sola seccion, siempre "Section1")

openpyxl no sabe nada de esto, asi que la parte se escribe a mano y el .xlsx
se reempaqueta: se agregan customXml/item1.xml, sus props y su relacion desde
xl/workbook.xml.rels, y se registran los tipos de contenido.
"""
import base64
import io
import os
import shutil
import struct
import zipfile
from xml.sax.saxutils import escape

NS_DATAMASHUP = "http://schemas.microsoft.com/DataMashup"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_OFFICE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"

REL_CUSTOM_XML = f"{NS_OFFICE_REL}/customXml"
REL_CUSTOM_XML_PROPS = f"{NS_OFFICE_REL}/customXmlProps"
CT_CUSTOM_XML_PROPS = ("application/vnd.openxmlformats-officedocument"
                       ".customXmlProperties+xml")

# GUID fijo: el archivo se regenera igual en cada corrida
ITEM_ID = "{4C1C7B23-9E5A-4A67-9B77-9E1D0E4A5F31}"

PACKAGE_CONTENT_TYPES = (
    '<?xml version="1.0" encoding="utf-8"?>'
    f'<Types xmlns="{NS_CT}">'
    '<Default Extension="m" ContentType="application/x-ms-m"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/Config/Package.xml" ContentType="application/json"/>'
    '</Types>'
)

PACKAGE_XML = (
    '<?xml version="1.0" encoding="utf-8"?>'
    f'<Package xmlns="{NS_DATAMASHUP}">'
    '<Version>2.145.1054.0</Version>'
    '<MinVersion>2.21.0.0</MinVersion>'
    '<Culture>es-AR</Culture>'
    '</Package>'
)

PERMISSIONS_XML = (
    '<?xml version="1.0" encoding="utf-8"?>'
    '<PermissionList xmlns:xsd="http://www.w3.org/2001/XMLSchema"'
    ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
    '<CanEvaluateFuturePackages>false</CanEvaluateFuturePackages>'
    '<FirewallEnabled>true</FirewallEnabled>'
    '</PermissionList>'
)

# fecha fija dentro del zip interno: dos corridas dan el mismo binario
FECHA_ZIP = (2026, 1, 1, 0, 0, 0)


def _zip_package_parts(seccion_m):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for nombre, contenido in (
            ("[Content_Types].xml", PACKAGE_CONTENT_TYPES),
            ("Config/Package.xml", PACKAGE_XML),
            ("Formulas/Section1.m", seccion_m),
        ):
            info = zipfile.ZipInfo(nombre, date_time=FECHA_ZIP)
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, contenido.encode("utf-8"))
    return buf.getvalue()


def _metadata_xml(consultas):
    """
    Un `<Item>` por consulta mas el item global `AllFormulas`.

    Las entradas usan el prefijo de tipo del formato: `l` para enteros y `s`
    para texto. Se declara cada consulta como "solo conexion": el destino real
    (el modelo de datos) se elige despues desde Excel.
    """
    items = [
        '<Item><ItemLocation><ItemType>AllFormulas</ItemType>'
        '<ItemPath/><ItemName/></ItemLocation><StableEntries/></Item>'
    ]
    for nombre, descripcion in consultas:
        items.append(
            '<Item><ItemLocation><ItemType>Formula</ItemType>'
            f'<ItemPath>Section1/{escape(nombre)}</ItemPath><ItemName/>'
            '</ItemLocation><StableEntries>'
            '<Entry Type="IsPrivate" Value="l0"/>'
            '<Entry Type="FillEnabled" Value="l0"/>'
            '<Entry Type="FillObjectType" Value="sConnectionOnly"/>'
            '<Entry Type="FillToDataModelEnabled" Value="l0"/>'
            '<Entry Type="ResultType" Value="sTable"/>'
            f'<Entry Type="Description" Value="s{escape(descripcion)}"/>'
            '</StableEntries></Item>'
        )
    return (
        '<?xml version="1.0" encoding="utf-8"?>'
        '<LocalPackageMetadataFile xmlns:xsd="http://www.w3.org/2001/XMLSchema"'
        ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        '<Items>' + "".join(items) + '</Items>'
        '</LocalPackageMetadataFile>'
    )


def _metadata_stream(consultas):
    xml = _metadata_xml(consultas).encode("utf-8")
    return (struct.pack("<i", 0)            # version del bloque de metadatos
            + struct.pack("<i", len(xml))   # longitud del XML
            + xml
            + struct.pack("<i", 0))         # contenido adicional: vacio


def construir_stream(seccion_m, consultas):
    """Devuelve el binario DataMashup completo, sin base64."""
    partes = _zip_package_parts(seccion_m)
    permisos = PERMISSIONS_XML.encode("utf-8")
    metadatos = _metadata_stream(consultas)
    bindings = b"\x00"   # la spec admite este valor en lugar del checksum

    salida = io.BytesIO()
    salida.write(struct.pack("<i", 0))              # Version
    for bloque in (partes, permisos, metadatos, bindings):
        salida.write(struct.pack("<i", len(bloque)))
        salida.write(bloque)
    return salida.getvalue()


def leer_stream(binario):
    """
    Vuelve a leer el binario que produjo `construir_stream`.

    Se usa como verificacion en el pipeline: si el parser recorre el stream
    entero y recupera el Section1.m original, el empaquetado esta bien armado.
    """
    pos = 0

    def entero():
        nonlocal pos
        (v,) = struct.unpack_from("<i", binario, pos)
        pos += 4
        return v

    def bloque():
        nonlocal pos
        largo = entero()
        datos = binario[pos:pos + largo]
        if len(datos) != largo:
            raise ValueError(f"bloque truncado: se esperaban {largo} bytes")
        pos += largo
        return datos

    version = entero()
    if version != 0:
        raise ValueError(f"version inesperada: {version}")
    partes, permisos, metadatos, bindings = bloque(), bloque(), bloque(), bloque()
    if pos != len(binario):
        raise ValueError(f"sobran {len(binario) - pos} bytes al final del stream")

    with zipfile.ZipFile(io.BytesIO(partes)) as z:
        nombres = z.namelist()
        seccion = z.read("Formulas/Section1.m").decode("utf-8")

    (mv,) = struct.unpack_from("<i", metadatos, 0)
    (mlen,) = struct.unpack_from("<i", metadatos, 4)
    metadata_xml = metadatos[8:8 + mlen].decode("utf-8")

    return {
        "partes": nombres,
        "seccion": seccion,
        "permisos": permisos.decode("utf-8"),
        "metadata_version": mv,
        "metadata_xml": metadata_xml,
        "bindings": bindings,
    }


def _item1_xml(binario):
    b64 = base64.b64encode(binario).decode("ascii")
    return ('<?xml version="1.0" encoding="utf-8"?>'
            f'<DataMashup xmlns="{NS_DATAMASHUP}">{b64}</DataMashup>')


ITEM_PROPS_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    f'<ds:datastoreItem ds:itemID="{ITEM_ID}"'
    ' xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">'
    f'<ds:schemaRefs><ds:schemaRef ds:uri="{NS_DATAMASHUP}"/></ds:schemaRefs>'
    '</ds:datastoreItem>'
)

ITEM_RELS_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    f'<Relationships xmlns="{NS_PKG_REL}">'
    f'<Relationship Id="rId1" Type="{REL_CUSTOM_XML_PROPS}" Target="itemProps1.xml"/>'
    '</Relationships>'
)


def _agregar_relacion(rels_xml, rel_id):
    marca = '</Relationships>'
    nueva = (f'<Relationship Id="{rel_id}" Type="{REL_CUSTOM_XML}"'
             ' Target="../customXml/item1.xml"/>')
    return rels_xml.replace(marca, nueva + marca)


def _proximo_rel_id(rels_xml):
    usados = set()
    for trozo in rels_xml.split('Id="')[1:]:
        usados.add(trozo.split('"')[0])
    i = 1
    while f"rId{i}" in usados:
        i += 1
    return f"rId{i}"


def _registrar_content_type(ct_xml):
    if "/customXml/itemProps1.xml" in ct_xml:
        return ct_xml
    override = ('<Override PartName="/customXml/itemProps1.xml"'
                f' ContentType="{CT_CUSTOM_XML_PROPS}"/>')
    return ct_xml.replace("</Types>", override + "</Types>")


def inyectar(ruta_xlsx, seccion_m, consultas):
    """Agrega la parte DataMashup a un .xlsx ya escrito."""
    binario = construir_stream(seccion_m, consultas)
    leer_stream(binario)   # falla temprano si el stream quedo mal armado

    tmp = ruta_xlsx + ".tmp"
    with zipfile.ZipFile(ruta_xlsx) as origen:
        nombres = origen.namelist()
        if "customXml/item1.xml" in nombres:
            raise ValueError("el libro ya tiene una parte customXml/item1.xml")
        rels = origen.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        rel_id = _proximo_rel_id(rels)

        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as destino:
            for info in origen.infolist():
                datos = origen.read(info.filename)
                if info.filename == "xl/_rels/workbook.xml.rels":
                    datos = _agregar_relacion(rels, rel_id).encode("utf-8")
                elif info.filename == "[Content_Types].xml":
                    datos = _registrar_content_type(
                        datos.decode("utf-8")).encode("utf-8")
                destino.writestr(info, datos)

            destino.writestr("customXml/item1.xml", _item1_xml(binario))
            destino.writestr("customXml/itemProps1.xml", ITEM_PROPS_XML)
            destino.writestr("customXml/_rels/item1.xml.rels", ITEM_RELS_XML)

    shutil.move(tmp, ruta_xlsx)
    return len(binario)


def verificar(ruta_xlsx):
    """Relee la parte desde el .xlsx terminado y devuelve el resumen."""
    with zipfile.ZipFile(ruta_xlsx) as z:
        item = z.read("customXml/item1.xml").decode("utf-8")
    ini = item.index(">", item.index("<DataMashup")) + 1
    fin = item.index("</DataMashup>")
    return leer_stream(base64.b64decode(item[ini:fin]))


if __name__ == "__main__":
    import consultas_m
    binario = construir_stream(consultas_m.SECCION, consultas_m.CONSULTAS)
    info = leer_stream(binario)
    print("bytes del stream :", len(binario))
    print("partes del paquete:", info["partes"])
    print("consultas         :", info["seccion"].count("shared "))
