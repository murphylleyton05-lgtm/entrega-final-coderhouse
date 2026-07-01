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
└── evidencias/                        # Capturas del test de estres (5 ejecuciones)
    ├── flujo_completo_n8n.jpeg
    ├── ejecucion_1_feliz.png
    ├── ejecucion_2_feliz.png
    ├── ejecucion_3_feliz.png
    ├── ejecucion_4_infeliz_email_vacio.png
    └── ejecucion_5_infeliz_api_error.png
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

En la carpeta [`evidencias/`](./evidencias) estan las capturas de las 5 ejecuciones:
- 3 ejecuciones exitosas (camino feliz).
- 2 ejecuciones del camino infeliz (email vacio y fallo de API) que verifican las rutas de error.

## Video demo

📹 Demo de 3 minutos: **[AGREGAR LINK DEL VIDEO AQUI]**

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
| Test de estres (5 ejecuciones) | Carpeta `evidencias/` |
| Filtro anti-bucle | IF valida email antes de procesar |
| Prompt dinamico | System + User prompt con variables del sistema |

---

**Lleyton Murphy** | lleyton-ia-page.netlify.app | linkedin.com/in/lleyton-murphy-3716093a3
