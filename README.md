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
├── n8n_flow_leads_vip.json            # JSON del flujo — importar en n8n
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

## Base de datos (Airtable)

Link de la base en modo lectura: [INSERTAR LINK AIRTABLE AQUI]

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
