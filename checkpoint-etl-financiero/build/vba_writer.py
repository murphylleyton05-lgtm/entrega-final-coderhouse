"""
vba_writer.py
=============
Genera un `vbaProject.bin` valido (proyecto VBA real, no texto en una celda)
para incrustar dentro de un libro .xlsm.

Implementa, sin dependencias externas:

  1. La compresion "Run Length Encoding" de VBA descripta en [MS-OVBA] 2.4.1
     (CompressedContainer / CompressedChunk / TokenSequence).
  2. Un escritor de contenedores OLE / CFB v3 (sectores de 512 bytes, FAT,
     MiniFAT, MiniStream y directorio con arbol binario ordenado).
  3. Los streams propios del proyecto VBA: `dir`, `_VBA_PROJECT`, `PROJECT`,
     `PROJECTwm` y un stream por modulo.

Los registros globales del stream `dir` (SysKind, LCID, code page y las
referencias a stdole / Office) se toman de un proyecto generado por Excel,
de modo que la cabecera del proyecto sea byte-a-byte equivalente a la que
produce la aplicacion real.
"""

from __future__ import annotations

import struct
from dataclasses import dataclass, field

# --------------------------------------------------------------------------
# 1. Compresion [MS-OVBA] 2.4.1
# --------------------------------------------------------------------------

# Cada chunk comprimido admite como maximo 4096 bytes de datos. Al codificar
# solo literales, 8 bytes de datos gastan 1 byte de FlagByte, con lo cual el
# maximo de bytes sin comprimir por chunk es 3640 -> 3640 + 455 = 4095 bytes.
_MAX_BYTES_POR_CHUNK = 3640


def _codificar_chunk(crudo: bytes) -> bytes:
    """Codifica un chunk usando unicamente tokens literales (FlagByte = 0x00)."""
    salida = bytearray()
    for i in range(0, len(crudo), 8):
        salida.append(0x00)          # los 8 tokens siguientes son literales
        salida += crudo[i:i + 8]
    return bytes(salida)


def comprimir(datos: bytes) -> bytes:
    """Devuelve un CompressedContainer con la firma 0x01 y sus chunks."""
    salida = bytearray(b"\x01")
    for i in range(0, len(datos), _MAX_BYTES_POR_CHUNK):
        chunk = _codificar_chunk(datos[i:i + _MAX_BYTES_POR_CHUNK])
        assert len(chunk) <= 4096, "chunk fuera de rango"
        # bit 15 = comprimido, bits 14..12 = firma 0b011, bits 11..0 = tamano-1
        cabecera = 0x8000 | 0x3000 | (len(chunk) - 1)
        salida += struct.pack("<H", cabecera) + chunk
    return bytes(salida)


def _bits_copy_token(diferencia: int) -> tuple[int, int]:
    """[MS-OVBA] 2.4.1.3.19.1: mascara de longitud y bits segun la posicion."""
    bits = 4
    while (1 << bits) < diferencia:
        bits += 1
    bits = max(4, min(12, bits))
    return (0xFFFF >> bits), bits


def descomprimir(datos: bytes) -> bytes:
    """Inverso de `comprimir`; se usa para verificar el ida y vuelta."""
    assert datos[0] == 0x01, "falta la firma 0x01"
    salida = bytearray()
    pos = 1
    while pos < len(datos):
        cabecera, = struct.unpack_from("<H", datos, pos)
        tam_chunk = (cabecera & 0x0FFF) + 3
        comprimido = (cabecera >> 15) & 1
        fin = pos + tam_chunk
        cur = pos + 2
        inicio_chunk = len(salida)
        if not comprimido:
            salida += datos[cur:cur + 4096]
            pos = fin
            continue
        while cur < fin:
            banderas = datos[cur]
            cur += 1
            for bit in range(8):
                if cur >= fin:
                    break
                if not (banderas >> bit) & 1:
                    salida.append(datos[cur])
                    cur += 1
                else:
                    token, = struct.unpack_from("<H", datos, cur)
                    cur += 2
                    mascara_long, bits = _bits_copy_token(len(salida) - inicio_chunk)
                    longitud = (token & mascara_long) + 3
                    desplazamiento = (token >> (16 - bits)) + 1
                    origen = len(salida) - desplazamiento
                    for k in range(longitud):
                        salida.append(salida[origen + k])
        pos = fin
    return bytes(salida)


# --------------------------------------------------------------------------
# 2. Escritor de contenedor OLE / CFB v3
# --------------------------------------------------------------------------

LIBRE = 0xFFFFFFFF
FIN_CADENA = 0xFFFFFFFE
SECTOR_FAT = 0xFFFFFFFD
SIN_STREAM = 0xFFFFFFFF

TAM_SECTOR = 512
TAM_MINI_SECTOR = 64
UMBRAL_MINI = 4096


@dataclass
class _Entrada:
    nombre: str
    tipo: int                      # 1 = storage, 2 = stream, 5 = raiz
    datos: bytes = b""
    hijos: list = field(default_factory=list)   # indices
    izq: int = SIN_STREAM
    der: int = SIN_STREAM
    hijo: int = SIN_STREAM
    sector_inicial: int = 0
    tamano: int = 0
    clsid: bytes = b"\x00" * 16


def _clave_orden(nombre: str):
    """Orden del directorio CFB: primero por longitud, luego por nombre en mayusculas."""
    return (len(nombre), nombre.upper())


class EscritorCFB:
    def __init__(self, clsid_raiz: bytes = b"\x00" * 16):
        self.entradas: list[_Entrada] = [_Entrada("Root Entry", 5, clsid=clsid_raiz)]

    # -- construccion del arbol -------------------------------------------
    def agregar_storage(self, nombre: str, padre: int = 0) -> int:
        idx = len(self.entradas)
        self.entradas.append(_Entrada(nombre, 1))
        self.entradas[padre].hijos.append(idx)
        return idx

    def agregar_stream(self, nombre: str, datos: bytes, padre: int = 0) -> int:
        idx = len(self.entradas)
        self.entradas.append(_Entrada(nombre, 2, datos=datos))
        self.entradas[padre].hijos.append(idx)
        return idx

    # -- serializacion -----------------------------------------------------
    def _armar_arbol(self, indices: list[int]) -> int:
        if not indices:
            return SIN_STREAM
        medio = len(indices) // 2
        raiz = indices[medio]
        self.entradas[raiz].izq = self._armar_arbol(indices[:medio])
        self.entradas[raiz].der = self._armar_arbol(indices[medio + 1:])
        return raiz

    def construir(self) -> bytes:
        # 2.1 enlazar los hijos de cada storage mediante un arbol binario ordenado
        for entrada in self.entradas:
            if entrada.tipo in (1, 5) and entrada.hijos:
                ordenados = sorted(entrada.hijos,
                                   key=lambda i: _clave_orden(self.entradas[i].nombre))
                entrada.hijo = self._armar_arbol(ordenados)

        # 2.2 repartir los streams entre el MiniStream y los sectores normales
        mini_stream = bytearray()
        mini_fat: list[int] = []
        streams_normales: list[tuple[_Entrada, int]] = []   # (entrada, nro de sectores)

        for entrada in self.entradas[1:]:
            if entrada.tipo != 2:
                continue
            entrada.tamano = len(entrada.datos)
            if entrada.tamano == 0:
                entrada.sector_inicial = FIN_CADENA
            elif entrada.tamano < UMBRAL_MINI:
                inicio = len(mini_stream) // TAM_MINI_SECTOR
                entrada.sector_inicial = inicio
                relleno = (-len(entrada.datos)) % TAM_MINI_SECTOR
                mini_stream += entrada.datos + b"\x00" * relleno
                cantidad = (len(entrada.datos) + TAM_MINI_SECTOR - 1) // TAM_MINI_SECTOR
                for k in range(cantidad):
                    mini_fat.append(inicio + k + 1 if k < cantidad - 1 else FIN_CADENA)
            else:
                cantidad = (entrada.tamano + TAM_SECTOR - 1) // TAM_SECTOR
                streams_normales.append((entrada, cantidad))

        # 2.3 tamanos de las estructuras auxiliares
        entradas_dir = len(self.entradas)
        sectores_dir = (entradas_dir * 128 + TAM_SECTOR - 1) // TAM_SECTOR
        sectores_minifat = ((len(mini_fat) * 4 + TAM_SECTOR - 1) // TAM_SECTOR) if mini_fat else 0
        sectores_mini_stream = (len(mini_stream) + TAM_SECTOR - 1) // TAM_SECTOR
        sectores_datos = sum(c for _, c in streams_normales)

        sectores_sin_fat = sectores_dir + sectores_minifat + sectores_mini_stream + sectores_datos

        # 2.4 la FAT se describe a si misma: iterar hasta el punto fijo
        entradas_por_sector_fat = TAM_SECTOR // 4
        sectores_fat = 1
        while True:
            total = sectores_sin_fat + sectores_fat
            necesarios = max(1, (total + entradas_por_sector_fat - 1) // entradas_por_sector_fat)
            if necesarios == sectores_fat:
                break
            sectores_fat = necesarios
        if sectores_fat > 109:
            raise ValueError("el proyecto excede los 109 sectores FAT (DIFAT no implementado)")

        # 2.5 asignacion de numeros de sector
        siguiente = 0
        nros_fat = list(range(siguiente, siguiente + sectores_fat)); siguiente += sectores_fat
        nros_dir = list(range(siguiente, siguiente + sectores_dir)); siguiente += sectores_dir
        nros_minifat = list(range(siguiente, siguiente + sectores_minifat)); siguiente += sectores_minifat
        nros_mini_stream = list(range(siguiente, siguiente + sectores_mini_stream)); siguiente += sectores_mini_stream

        total_sectores = siguiente + sectores_datos
        fat = [LIBRE] * (sectores_fat * entradas_por_sector_fat)

        def encadenar(nros: list[int]):
            for i, s in enumerate(nros):
                fat[s] = nros[i + 1] if i < len(nros) - 1 else FIN_CADENA

        for s in nros_fat:
            fat[s] = SECTOR_FAT
        encadenar(nros_dir)
        encadenar(nros_minifat)
        encadenar(nros_mini_stream)

        cuerpos: list[tuple[int, bytes]] = []   # (primer sector, datos ya alineados)
        for entrada, cantidad in streams_normales:
            nros = list(range(siguiente, siguiente + cantidad))
            siguiente += cantidad
            encadenar(nros)
            entrada.sector_inicial = nros[0]
            relleno = (-len(entrada.datos)) % TAM_SECTOR
            cuerpos.append((nros[0], entrada.datos + b"\x00" * relleno))

        # 2.6 entrada raiz: apunta al MiniStream
        raiz = self.entradas[0]
        raiz.sector_inicial = nros_mini_stream[0] if nros_mini_stream else FIN_CADENA
        raiz.tamano = len(mini_stream)

        # 2.7 directorio
        dir_bytes = bytearray()
        for entrada in self.entradas:
            nombre_utf16 = entrada.nombre.encode("utf-16-le") + b"\x00\x00"
            if len(nombre_utf16) > 64:
                raise ValueError(f"nombre demasiado largo: {entrada.nombre}")
            dir_bytes += nombre_utf16.ljust(64, b"\x00")
            dir_bytes += struct.pack("<H", len(nombre_utf16))
            dir_bytes += struct.pack("<BB", entrada.tipo, 1)       # 1 = nodo negro
            dir_bytes += struct.pack("<III", entrada.izq, entrada.der, entrada.hijo)
            dir_bytes += entrada.clsid
            dir_bytes += struct.pack("<I", 0)                       # state bits
            dir_bytes += b"\x00" * 16                               # creacion / modificacion
            dir_bytes += struct.pack("<I", entrada.sector_inicial)
            dir_bytes += struct.pack("<Q", entrada.tamano)
        relleno = (-len(dir_bytes)) % TAM_SECTOR
        dir_bytes += b"\xFF" * 0 + b"\x00" * relleno

        # 2.8 cabecera
        cabecera = bytearray(512)
        cabecera[0:8] = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"
        cabecera[24:26] = struct.pack("<H", 0x003E)     # version menor
        cabecera[26:28] = struct.pack("<H", 0x0003)     # version mayor (sectores de 512)
        cabecera[28:30] = struct.pack("<H", 0xFFFE)     # little endian
        cabecera[30:32] = struct.pack("<H", 9)          # 2^9 = 512
        cabecera[32:34] = struct.pack("<H", 6)          # 2^6 = 64
        cabecera[40:44] = struct.pack("<I", 0)          # cantidad de sectores de directorio (v3)
        cabecera[44:48] = struct.pack("<I", sectores_fat)
        cabecera[48:52] = struct.pack("<I", nros_dir[0])
        cabecera[52:56] = struct.pack("<I", 0)
        cabecera[56:60] = struct.pack("<I", UMBRAL_MINI)
        cabecera[60:64] = struct.pack("<I", nros_minifat[0] if nros_minifat else FIN_CADENA)
        cabecera[64:68] = struct.pack("<I", sectores_minifat)
        cabecera[68:72] = struct.pack("<I", FIN_CADENA)  # primer sector DIFAT
        cabecera[72:76] = struct.pack("<I", 0)           # cantidad de sectores DIFAT
        difat = [nros_fat[i] if i < len(nros_fat) else LIBRE for i in range(109)]
        cabecera[76:512] = struct.pack("<109I", *difat)

        # 2.9 ensamblado final
        salida = bytearray(cabecera)
        salida += b"\x00" * (total_sectores * TAM_SECTOR)

        def escribir_sector(nro: int, datos: bytes):
            ini = 512 + nro * TAM_SECTOR
            salida[ini:ini + len(datos)] = datos

        fat_bytes = struct.pack(f"<{len(fat)}I", *fat)
        for i, s in enumerate(nros_fat):
            escribir_sector(s, fat_bytes[i * TAM_SECTOR:(i + 1) * TAM_SECTOR])
        for i, s in enumerate(nros_dir):
            escribir_sector(s, bytes(dir_bytes[i * TAM_SECTOR:(i + 1) * TAM_SECTOR]))
        if mini_fat:
            mini_fat_bytes = struct.pack(f"<{len(mini_fat)}I", *mini_fat)
            mini_fat_bytes += b"\xFF" * ((-len(mini_fat_bytes)) % TAM_SECTOR)
            for i, s in enumerate(nros_minifat):
                escribir_sector(s, mini_fat_bytes[i * TAM_SECTOR:(i + 1) * TAM_SECTOR])
        mini_stream_bytes = bytes(mini_stream)
        for i, s in enumerate(nros_mini_stream):
            escribir_sector(s, mini_stream_bytes[i * TAM_SECTOR:(i + 1) * TAM_SECTOR])
        for primer_sector, datos in cuerpos:
            for i in range(0, len(datos), TAM_SECTOR):
                escribir_sector(primer_sector + i // TAM_SECTOR, datos[i:i + TAM_SECTOR])

        return bytes(salida)


# --------------------------------------------------------------------------
# 3. Streams propios del proyecto VBA
# --------------------------------------------------------------------------

GUID_LIBRO = "{00020819-0000-0000-C000-000000000046}"
GUID_HOJA = "{00020820-0000-0000-C000-000000000046}"

CODEPAGE = 1252          # 0x04E4, igual que el proyecto de referencia


@dataclass
class Modulo:
    nombre: str
    codigo: str
    tipo: str = "procedural"       # "procedural" | "document"
    guid_base: str | None = None   # solo para modulos de documento


def _reg(id_registro: int, cuerpo: bytes) -> bytes:
    return struct.pack("<HI", id_registro, len(cuerpo)) + cuerpo


def _mbcs(texto: str) -> bytes:
    return texto.encode(f"cp{CODEPAGE}")


def _utf16(texto: str) -> bytes:
    return texto.encode("utf-16-le")


def construir_dir(nombre_proyecto: str, modulos: list[Modulo]) -> bytes:
    """Arma el stream `dir` sin comprimir ([MS-OVBA] 2.3.4.2)."""
    d = bytearray()

    # --- PROJECTINFORMATION ---
    d += _reg(0x0001, struct.pack("<I", 1))            # SysKind: 32 bits
    d += _reg(0x0002, struct.pack("<I", 0x0409))       # LCID
    d += _reg(0x0014, struct.pack("<I", 0x0409))       # LCIDINVOKE
    d += _reg(0x0003, struct.pack("<H", CODEPAGE))     # CodePage
    d += _reg(0x0004, _mbcs(nombre_proyecto))          # Name
    d += _reg(0x0005, b"")                             # DocString
    d += _reg(0x0040, b"")                             # DocString unicode
    d += _reg(0x0006, b"")                             # HelpFilePath
    d += _reg(0x003D, b"")                             # HelpFilePath 2
    d += _reg(0x0007, struct.pack("<I", 0))            # HelpContext
    d += _reg(0x0008, struct.pack("<I", 0))            # LibFlags
    # PROJECTVERSION: el campo de tamano es un reservado fijo (0x00000004)
    d += struct.pack("<HI", 0x0009, 4)
    d += struct.pack("<IH", 1382489937, 32)            # VersionMajor / VersionMinor
    d += _reg(0x000C, b"")                             # Constants
    d += _reg(0x003C, b"")                             # Constants unicode

    # --- PROJECTREFERENCES (identicas a las que genera Excel) ---
    for nombre, libid in (
        ("stdole",
         r"*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\WINDOWS\system32\stdole2.tlb#OLE Automation"),
        ("Office",
         r"*\G{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.0#0#C:\Program Files\Common Files"
         r"\Microsoft Shared\OFFICE12\MSO.DLL#Microsoft Office 12.0 Object Library"),
    ):
        d += _reg(0x0016, _mbcs(nombre))               # REFERENCENAME
        d += _reg(0x003E, _utf16(nombre))              # REFERENCENAME unicode
        libid_bytes = _mbcs(libid)
        cuerpo = struct.pack("<I", len(libid_bytes)) + libid_bytes + struct.pack("<IH", 0, 0)
        d += struct.pack("<HI", 0x000D, len(cuerpo)) + cuerpo   # REFERENCEREGISTERED

    # --- PROJECTMODULES ---
    d += _reg(0x000F, struct.pack("<H", len(modulos)))
    d += _reg(0x0013, struct.pack("<H", 0xFFFF))       # PROJECTCOOKIE

    for modulo in modulos:
        d += _reg(0x0019, _mbcs(modulo.nombre))        # MODULENAME
        d += _reg(0x0047, _utf16(modulo.nombre))       # MODULENAME unicode
        d += _reg(0x001A, _mbcs(modulo.nombre))        # MODULESTREAMNAME
        d += _reg(0x0032, _utf16(modulo.nombre))       # MODULESTREAMNAME unicode
        d += _reg(0x001C, b"")                         # MODULEDOCSTRING
        d += _reg(0x0048, b"")                         # MODULEDOCSTRING unicode
        d += _reg(0x0031, struct.pack("<I", 0))        # MODULEOFFSET: el stream es solo codigo
        d += _reg(0x001E, struct.pack("<I", 0))        # MODULEHELPCONTEXT
        d += _reg(0x002C, struct.pack("<H", 0xFFFF))   # MODULECOOKIE
        d += _reg(0x0022 if modulo.tipo == "document" else 0x0021, b"")   # MODULETYPE
        d += _reg(0x002B, b"")                         # Terminador del modulo

    d += _reg(0x0010, struct.pack("<I", 0))            # Terminador del stream
    return bytes(d)


def construir_project(nombre_proyecto: str, modulos: list[Modulo]) -> bytes:
    """Stream `PROJECT`: el indice en texto plano que lee el editor de VBA."""
    lineas = ['ID="{8D807122-0657-42C8-BC6F-4B5FD08031C9}"']
    for modulo in modulos:
        if modulo.tipo == "document":
            lineas.append(f"Document={modulo.nombre}/&H00000000")
        else:
            lineas.append(f"Module={modulo.nombre}")
    lineas += [
        f'Name="{nombre_proyecto}"',
        'HelpContextID="0"',
        'VersionCompatible32="393222000"',
        'CMG="DBD966ECE4F0E4F0E4F0E4F0"',
        'DPB="5557E86E18E919E919E9"',
        'GC="CFCD72F0961011111111EE"',
        "",
        "[Host Extender Info]",
        "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000",
        "",
        "[Workspace]",
    ]
    for modulo in modulos:
        sufijo = "C" if modulo.tipo == "document" else ""
        lineas.append(f"{modulo.nombre}=0, 0, 0, 0, {sufijo}")
    return ("\r\n".join(lineas) + "\r\n").encode(f"cp{CODEPAGE}")


def construir_projectwm(modulos: list[Modulo]) -> bytes:
    """Stream `PROJECTwm`: pares nombre MBCS / nombre Unicode."""
    salida = bytearray()
    for modulo in modulos:
        salida += _mbcs(modulo.nombre) + b"\x00"
        salida += _utf16(modulo.nombre) + b"\x00\x00"
    salida += b"\x00\x00"
    return bytes(salida)


def _encabezado_modulo(modulo: Modulo) -> str:
    """Lineas `Attribute` que VBA guarda al principio de cada modulo."""
    lineas = [f'Attribute VB_Name = "{modulo.nombre}"']
    if modulo.tipo == "document":
        lineas += [
            f'Attribute VB_Base = "0{modulo.guid_base}"',
            "Attribute VB_GlobalNameSpace = False",
            "Attribute VB_Creatable = False",
            "Attribute VB_PredeclaredId = True",
            "Attribute VB_Exposed = True",
            "Attribute VB_TemplateDerived = False",
            "Attribute VB_Customizable = True",
        ]
    return "\r\n".join(lineas) + "\r\n"


def construir_vba_project(modulos: list[Modulo], nombre_proyecto: str = "VBAProject") -> bytes:
    """Devuelve el contenido completo de `xl/vbaProject.bin`."""
    cfb = EscritorCFB()

    cfb.agregar_stream("PROJECT", construir_project(nombre_proyecto, modulos))
    cfb.agregar_stream("PROJECTwm", construir_projectwm(modulos))

    vba = cfb.agregar_storage("VBA")

    # _VBA_PROJECT: se declara una version inexistente (0xFFFF) para que el host
    # descarte la cache de compilacion y recompile a partir del codigo fuente.
    cfb.agregar_stream("_VBA_PROJECT", b"\xCC\x61\xFF\xFF\x00\x00\x00", padre=vba)
    cfb.agregar_stream("dir", comprimir(construir_dir(nombre_proyecto, modulos)), padre=vba)

    for modulo in modulos:
        fuente = _encabezado_modulo(modulo) + modulo.codigo.replace("\n", "\r\n")
        cfb.agregar_stream(modulo.nombre, comprimir(fuente.encode(f"cp{CODEPAGE}")), padre=vba)

    return cfb.construir()
