"""Guiones reales y factualmente correctos para los temas del banco offline.

Sirven cuando no hay IA disponible: el video igual EXPLICA el dato de verdad,
en vez de relleno vacío. Con IA (Gemini/Claude) se generan temas ilimitados.
"""
from __future__ import annotations

import unicodedata


def _norm(s: str) -> str:
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode().lower()
    return "".join(c for c in s if c.isalnum() or c == " ").strip()


# Cada guion: gancho + explicación real + cierre. ~80-110 palabras.
_BANCO = {
    "por que los pulpos tienen tres corazones y sangre azul": (
        "¿Sabías que el pulpo tiene tres corazones? Dos bombean sangre a las "
        "branquias y el tercero al resto del cuerpo. Pero hay algo más raro: su "
        "sangre es azul. En vez de hemoglobina con hierro, como la nuestra, usa "
        "hemocianina, que tiene cobre. Ese cobre es mucho más eficiente para "
        "transportar oxígeno en aguas frías y con poco oxígeno del fondo del mar. "
        "Y el dato final: cuando el pulpo nada, su corazón principal deja de "
        "latir. Por eso prefiere caminar, porque nadar lo agota. Seguime para más."
    ),
    "el unico metal liquido a temperatura ambiente ademas del mercurio": (
        "El mercurio es famoso por ser el único metal líquido a temperatura "
        "ambiente. Pero hay otro que se derrite con solo agarrarlo: el galio. "
        "El galio es sólido a temperatura ambiente, pero su punto de fusión es de "
        "apenas veintinueve grados. Como tu mano está a treinta y seis, si "
        "sostenés un trozo, se derrite en tu palma como si fuera hielo. Es un "
        "metal real, brillante, que podés fundir con el calor del cuerpo, y "
        "cuando lo soltás vuelve a endurecerse. Seguime para un dato nuevo cada día."
    ),
    "por que la miel nunca se echa a perder": (
        "¿Sabías que la miel nunca se echa a perder? Encontraron miel en tumbas "
        "egipcias de hace más de tres mil años, y todavía era comestible. El "
        "secreto está en tres cosas. Primero, casi no tiene agua, así que las "
        "bacterias no pueden sobrevivir. Segundo, es muy ácida. Y tercero, las "
        "abejas le agregan una enzima que produce peróxido de hidrógeno, el mismo "
        "del agua oxigenada. Entre la falta de agua, el ácido y ese antiséptico "
        "natural, ningún microbio prospera. Por eso es el único alimento que "
        "prácticamente no caduca. Seguime para más datos así."
    ),
    "el pais que tiene mas husos horarios del mundo": (
        "¿Qué país abarca más husos horarios? La respuesta sorprende: es Francia. "
        "Aunque el territorio europeo entra en una sola zona, Francia tiene islas "
        "y territorios repartidos por todo el planeta, desde el Pacífico hasta el "
        "Caribe. Sumando todos, cubre doce husos horarios. Rusia, que parece la "
        "obvia por su tamaño, tiene once. Estados Unidos, con sus territorios, "
        "también anda cerca. Pero el campeón absoluto es Francia, gracias a su "
        "imperio colonial repartido por los océanos. Seguí el canal para más."
    ),
    "por que los flamencos son rosados y no nacen asi": (
        "Los flamencos no nacen rosados. Cuando salen del huevo son grises. Su "
        "color viene de lo que comen: algas y pequeños crustáceos, como camarones "
        "diminutos, que contienen unos pigmentos llamados carotenoides. Es el "
        "mismo tipo de pigmento que hace naranja a la zanahoria. A medida que el "
        "flamenco come, esos pigmentos se acumulan en sus plumas y lo van tiñendo "
        "de rosa. Tanto es así que un flamenco mal alimentado se vuelve pálido. "
        "Literalmente, son del color de su comida. Seguime para más curiosidades."
    ),
    "cuanto tiempo tarda la luz del sol en llegar a la tierra": (
        "Cuando mirás el Sol, no lo estás viendo como es ahora, sino como era hace "
        "ocho minutos. La luz viaja a trescientos mil kilómetros por segundo, "
        "rapidísimo, pero el Sol está a ciento cincuenta millones de kilómetros. "
        "Hacé la cuenta y esa luz tarda unos ocho minutos y veinte segundos en "
        "llegar. Eso significa que si el Sol se apagara de golpe, seguiríamos "
        "viéndolo brillar más de ocho minutos antes de enterarnos. Mirar el cielo "
        "es, en realidad, mirar el pasado. Seguí el canal para más."
    ),
    "el animal que puede sobrevivir en el espacio el tardigrado": (
        "Existe un animal que sobrevive en el espacio exterior, y mide menos de un "
        "milímetro: el tardígrado. En dos mil siete, los científicos los mandaron "
        "al vacío del espacio, sin aire, con radiación letal y frío extremo. "
        "Muchos volvieron vivos. Su truco es un estado llamado criptobiosis: se "
        "secan casi por completo, frenan su metabolismo a casi cero y se vuelven "
        "una cápsula indestructible. Resisten temperaturas de menos doscientos "
        "grados y presiones que aplastarían a cualquier otro ser. Son casi "
        "inmortales. Seguime para más datos increíbles."
    ),
    "por que el corazon de la ballena azul es del tamano de un auto": (
        "El corazón de la ballena azul es tan grande como un auto pequeño y pesa "
        "unos ciento cincuenta kilos. Tiene sentido: la ballena azul es el animal "
        "más grande que existió jamás, más que cualquier dinosaurio, y llega a "
        "treinta metros. Para mover sangre por semejante cuerpo, necesita una "
        "bomba enorme. Cada latido impulsa cientos de litros de sangre, y sus "
        "arterias son tan anchas que un humano podría gatear por dentro. Late tan "
        "fuerte que se puede detectar a tres kilómetros. Seguí para más."
    ),
    "el lugar mas seco de la tierra donde nunca llovio": (
        "Hay un lugar en la Tierra donde nunca, jamás, se registró lluvia: los "
        "Valles Secos de la Antártida. Suena raro que el lugar más seco del "
        "planeta esté en el continente helado, pero es así. Vientos poderosísimos, "
        "que superan los trescientos kilómetros por hora, evaporan toda la "
        "humedad antes de que toque el suelo. Se estima que en algunas zonas no "
        "llueve hace dos millones de años. Es tan parecido a Marte que la NASA "
        "prueba ahí sus robots. Seguime para más lugares imposibles."
    ),
    "por que los bananos son bayas pero las frutillas no": (
        "Preparate, porque la botánica es un caos. La banana es, técnicamente, una "
        "baya. Y la frutilla no lo es. Para la ciencia, una baya es una fruta que "
        "nace de una sola flor con un ovario, y que lleva las semillas dentro de "
        "la pulpa. La banana cumple. La frutilla, en cambio, nace de una flor con "
        "muchos ovarios, y sus semillas, esos puntitos de afuera, son en realidad "
        "las frutas verdaderas. Lo rojo y jugoso es solo el soporte. Todo lo que "
        "creías estaba al revés. Seguí el canal."
    ),
    "cuantas veces late tu corazon en toda tu vida": (
        "¿Cuántas veces late tu corazón en toda tu vida? El corazón late unas "
        "setenta veces por minuto. Eso son cien mil latidos por día, y unos "
        "treinta y cinco millones por año. Si vivís ochenta años, tu corazón "
        "habrá latido cerca de tres mil millones de veces, sin descansar ni un "
        "segundo. Es el músculo más incansable del cuerpo. Y hay un patrón "
        "curioso en la naturaleza: casi todos los mamíferos, del ratón al "
        "elefante, viven aproximadamente la misma cantidad total de latidos. "
        "Seguime para más datos del cuerpo."
    ),
    "por que saturno flotaria en una pileta gigante": (
        "Si existiera una pileta lo bastante grande, Saturno flotaría en el agua. "
        "Aunque es un planeta gigante, está hecho casi todo de hidrógeno y helio, "
        "los gases más livianos del universo. Su densidad es menor que la del "
        "agua: si pudieras meterlo en un océano descomunal, no se hundiría, "
        "flotaría como un corcho. Es el único planeta del sistema solar al que le "
        "pasa esto. Los demás, incluida la Tierra, se irían al fondo de una. "
        "Seguí el canal para más datos del espacio."
    ),
    "por que el rayo es mas caliente que la superficie del sol": (
        "Un rayo es más caliente que la superficie del Sol. La superficie solar "
        "está a unos cinco mil quinientos grados. Un rayo, en el instante en que "
        "cae, alcanza treinta mil grados, cinco veces más. Pasa porque la "
        "electricidad atraviesa el aire en millonésimas de segundo y lo calienta "
        "de golpe a lo bestia. Ese calor repentino expande el aire tan rápido que "
        "genera una onda de choque: eso es el trueno que escuchás. Por un "
        "instante, un pedacito de cielo es más caliente que el Sol. Seguime para más."
    ),
}


def buscar(topic: str) -> str | None:
    """Devuelve el guion real si el tema está en el banco, o None."""
    n = _norm(topic)
    if n in _BANCO:
        return _BANCO[n]
    # Coincidencia parcial (por si el tema viene con leves variaciones).
    for clave, guion in _BANCO.items():
        if n and (n in clave or clave in n):
            return guion
    return None
