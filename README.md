# Ecosistema de Automatizacion IA — Clasificacion de Leads VIP

**Entrega Final | CoderCamp IA Automation | Lleyton Murphy**

---

## Descripcion del proyecto

Sistema de automatizacion end-to-end que clasifica leads entrantes con IA (Claude), los registra en Airtable, envia una notificacion de validacion al equipo por Slack (Human-in-the-loop) y envia la propuesta personalizada por Gmail **solo si un humano la aprueba**.

## Stack tecnologico

| Componente | Tecnologia |
|------------|------------|
| Orquestador | n8n (self-hosted / cloud) |
| Base de datos | Airtable |
| Motor de IA | Claude (Anthropic) |
| Notificacion HITL | Slack |
| Canal de salida | Gmail |

## Archivos del repositorio

```
/
├── README.md                          # Este archivo
├── EntregaFinal_Murphy_Lleyton.pdf    # Diagrama de arquitectura + documentacion completa
├── n8n_flow_leads_vip.json            # Flujo principal — importar en n8n
├── n8n_flow_gmail_triage.json         # Flujo bonus: triage de correos entrantes con IA
└── evidencias/                        # Evidencia visual del flujo
    ├── README.md                      # Escenarios del test de estres
    └── flujo_completo_n8n.jpeg        # Flujo completo desplegado en n8n
```

## Flujo del sistema

```
Formulario web
      ↓
[Webhook n8n] → [Filtro email] → [Airtable: crear lead]
                                          ↓
                              [Claude API: clasificar + propuesta]
                                          ↓
                              [Airtable: actualizar score/estado]
                                          ↓
                              [Filtro: solo VIP?]
                               ↙              ↘
                    [Slack HITL]          [Archivar no-VIP]
                          ↓
                  [Esperar aprobacion humana]
                    ↙         ↘
             [Gmail]        [Rechazar]
                ↓
        [Airtable: estado final]
```

## Base de datos (Airtable)

Base en modo lectura: https://airtable.com/appsu1jiEYix3PWMd/shrY3PoiBHvqbvwxI

Estructura relacional:
- **Leads** (tabla principal): registro de cada prospecto con campos de estado (`Pendiente`, `Procesado por IA`, `Aprobado y Enviado`, `Rechazado`).
- **Empresas** (tabla vinculada): datos de la empresa de cada lead, relacionada con Leads para evitar datos aislados.

## Evidencias (Test de estres)

En la carpeta [`evidencias/`](./evidencias) esta la evidencia del flujo desplegado en n8n.
Los 5 escenarios del test de estres (3 de camino feliz + 2 de camino infeliz: email vacio y
fallo de API) estan documentados en el PDF de arquitectura, seccion 6.

## Como importar el flujo en n8n

1. Abri n8n → Menu → Import from file
2. Selecciona `n8n_flow_leads_vip.json`
3. Configura las credenciales en cada nodo (Airtable, Claude, Slack, Gmail)
4. Activa el flujo

## Mapeo con las consignas

| Requisito de la consigna | Donde se cumple |
|--------------------------|-----------------|
| Caso de uso con lenguaje natural | Clasificacion de leads VIP con Claude |
| Base de datos con campos de estado y relaciones | Airtable: tablas Leads + Empresas |
| Trigger inteligente | Webhook (no polling) |
| Motor de IA con Max Tokens y mapeo de respuesta | Claude, `max_tokens` limitado, se mapea `content[0].text` |
| Gestion de errores (resiliencia) | Nodo Error Trigger + rutas IF + registro en `Error_log` |
| Human-in-the-loop | Notificacion Slack + espera de aprobacion antes de enviar |
| Salida multicanal | Slack (validacion) + Gmail (propuesta) |
| Test de estres (5 escenarios) | Documentado en PDF seccion 6 + `evidencias/` |
| Filtro anti-bucle | IF valida email antes de procesar |
| Prompt dinamico | System + User prompt con variables del sistema |

## Otras entregas en este repositorio

Ademas de la entrega final de CoderCamp IA Automation (todo lo de arriba), el
repositorio guarda los checkpoints de **Coderhouse — Data Analyst**:

| Carpeta | Entrega |
|---------|---------|
| [`RetailPro/`](./RetailPro) | Proyecto RetailPro — M3: script SQL de la base `Ventas_Tech_DB` (DDL, restricciones de integridad y carga de datos). M4: consultas de negocio con funciones de agregacion |
| [`modulo9_dashboard_final/`](./modulo9_dashboard_final) | Modulo 9 — capa de visualizacion: dashboard en Excel con segmentadores, cascada, dispersion y narrativa, mas el registro de prompts de IA |

## Extra: el cerebro de Claude en 3D

[`cerebro_claude_3d.html`](./cerebro_claude_3d.html) es un modelo 3D interactivo (three.js) de un "cerebro"
que representa, como metafora, como piensa Claude: 7 regiones (razonamiento, atencion, lenguaje,
vision, codigo, valores y memoria), senales que viajan entre neuronas y un modo **Transformer**
que muestra lo que Claude es de verdad: capas de atencion causal.

- **Abrirlo:** doble clic en el archivo. Funciona sin internet gracias a las copias de three.js en `vendor/three/` (licencia MIT).
- **Publicarlo en GitHub Pages:** el workflow `.github/workflows/pages.yml` lo publica al hacer push a `main`.
  Hay que activarlo una vez en *Settings → Pages → Source: GitHub Actions*.
- **Funciones:** rotar y hacer zoom, tour guiado, vista separada, corte coronal, temas de color,
  modo sueno, sonido de sinapsis, voz, historial y captura en PNG.
- **Respuestas reales de Claude:** abierto dentro de claude.ai (como artifact), el cuadro "Pensar" responde con Claude,
  recuerda la conversacion, mira imagenes adjuntas, se sincroniza con otras personas conectadas y guarda un muro de
  pensamientos compartido. Fuera de claude.ai (archivo local o GitHub Pages) funciona en modo simulado.

### Cerebros conectados: Claude + ChatGPT

[`cerebros_conectados.html`](./cerebros_conectados.html) pone el cerebro de Claude junto al de ChatGPT ("Luna")
unidos por un puente. Haces una pregunta y cada IA muestra que regiones usa y responde; en modo **Debate**
ChatGPT responde primero, su respuesta cruza el puente y Claude le contesta.

En el recuadro "Conexion" eliges el proveedor:

| Proveedor | Costo | Cerebro derecho | Cerebro izquierdo |
|-----------|-------|-----------------|-------------------|
| **Ollama** (local) | Gratis, sin internet | `gpt-oss:20b` (modelo abierto de OpenAI) | `llama3.2` (en lugar de Claude) |
| **Groq** | Gratis con limite de uso | `openai/gpt-oss-20b` | `llama-3.3-70b-versatile` (en lugar de Claude) |
| **n8n** | APIs pagas de OpenAI y Anthropic | ChatGPT | Claude |
| **claude.ai** (artifact) | Gratis con tu plan | simulado | Claude real |

**Ollama, paso a paso:**

1. Instala Ollama desde [ollama.com](https://ollama.com).
2. En una terminal: `ollama pull gpt-oss:20b` y `ollama pull llama3.2` (gpt-oss:20b necesita unos 16 GB de RAM).
3. Si abres la pagina como archivo o desde GitHub Pages, inicia Ollama permitiendo el acceso del navegador:
   `OLLAMA_ORIGINS="*" ollama serve` (en Windows: `set OLLAMA_ORIGINS=*` y luego `ollama serve`).
4. En la pagina elige "Ollama", pulsa "Guardar y probar" y pregunta.

**Groq:** crea una clave gratis en [console.groq.com](https://console.groq.com), pegala en la pagina y pulsa
"Guardar y probar". La clave queda solo en tu navegador: no publiques la pagina con tu clave adentro.

**n8n:** importa [`n8n_flow_cerebros_conectados.json`](./n8n_flow_cerebros_conectados.json), define
`OPENAI_API_KEY` y `ANTHROPIC_API_KEY` (opcionales: `OPENAI_MODEL`, `CLAUDE_MODEL`), activa el flujo y pega la URL
de produccion del webhook. Las claves quedan en n8n; la pagina nunca las ve.

Claude no tiene API gratis: con Ollama o Groq el cerebro izquierdo es otro modelo abierto, y la pagina lo dice.

---

**Lleyton Murphy** | lleyton-ia-page.netlify.app | linkedin.com/in/lleyton-murphy-3716093a3
