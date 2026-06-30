# Ecosistema de Automatizacion IA — Clasificacion de Leads VIP

**Entrega Final | CoderCamp IA Automation | Lleyton Murphy**

---

## Descripcion del proyecto

Sistema de automatizacion end-to-end que clasifica leads entrantes con IA (Claude), los registra en Airtable, envia una notificacion de validacion al equipo por Slack (HITL) y envia la propuesta personalizada por Gmail solo si un humano lo aprueba.

## Stack tecnologico

| Componente | Tecnologia |
|------------|------------|
| Orquestador | n8n |
| Base de datos | Airtable |
| Motor de IA | Claude (Anthropic) — claude-sonnet-4-6 |
| Notificacion HITL | Slack |
| Canal de salida | Gmail |

## Archivos del repositorio

```
/
├── README.md                          # Este archivo
├── EntregaFinal_Murphy_Lleyton.pdf    # Diagrama de arquitectura completo
├── n8n_flow_leads_vip.json            # JSON del flujo principal (leads VIP) — importar en n8n
├── n8n_flow_gmail_triage.json         # JSON del flujo de triage de Gmail entrante — importar en n8n
└── screenshots/                       # Evidencias de ejecucion
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
                  [Esperar aprobacion]
                    ↙         ↘
             [Gmail]        [Rechazar]
                ↓
        [Airtable: estado final]
```

## Flujo complementario: Triage de Gmail entrante

Flujo de **entrada** que complementa al canal de salida del sistema principal. Clasifica los correos que llegan a la bandeja, los etiqueta y deja preparado un borrador de respuesta para revision humana (no envia automaticamente).

```
[Gmail Trigger: correo nuevo no leido]
            ↓
   [Filtro: remitente valido]
            ↓
[Claude API: clasificar + redactar borrador]
            ↓
   [Funcion: parsear JSON]
            ↓
   [Switch: por categoria]
   ↙      ↓       ↓      ↘
Comercial Soporte Spam  Personal
   ↓        ↓      ↓       ↓
[etiqueta][etiqueta][etiqueta][etiqueta]
   ↓        ↓
[borrador][borrador]
```

Categorias: `Comercial`, `Soporte`, `Spam`, `Personal`. Solo Comercial y Soporte generan borrador de respuesta. Antes de activar, crea las etiquetas en Gmail y reemplaza los `labelIds` (`Label_Comercial`, `Label_Soporte`, `Label_Spam`, `Label_Personal`) por los IDs reales de tu cuenta.

## Base de datos (Airtable)

Link de la base en modo lectura: https://airtable.com/appsu1jiEYix3PWMd/shrY3PoiBHvqbvwxI

## Variables de entorno requeridas

```
ANTHROPIC_API_KEY=sk-ant-...
AIRTABLE_BASE_ID=app...
AIRTABLE_TOKEN=pat...
SLACK_BOT_TOKEN=xoxb-...
GMAIL_FROM=tu@gmail.com
```

## Como importar el flujo en n8n

1. Abri n8n → Menu → Import from file
2. Selecciona `n8n_flow_leads_vip.json`
3. Configura las credenciales en cada nodo (Airtable, Claude, Slack, Gmail)
4. Activa el flujo

---

**Lleyton Murphy** | lleyton-ia-page.netlify.app | linkedin.com/in/lleyton-murphy-3716093a3
