# Evidencias de ejecución

Esta carpeta contiene las capturas del **test de estrés** (mínimo 5 ejecuciones,
incluyendo camino feliz y camino infeliz) exigido por la consigna.

## Archivos

| Archivo | Qué muestra |
|---------|-------------|
| `flujo_completo_n8n.jpeg` | El flujo completo armado en n8n cloud |
| `ejecucion_1_feliz.png` | Ejecución exitosa — lead VIP clasificado y enviado |
| `ejecucion_2_feliz.png` | Ejecución exitosa — segundo lead |
| `ejecucion_3_feliz.png` | Ejecución exitosa — registro en Airtable actualizado |
| `ejecucion_4_infeliz_email_vacio.png` | Camino infeliz — email vacío, el filtro detiene el flujo |
| `ejecucion_5_infeliz_api_error.png` | Camino infeliz — fallo de API, se registra en Error_log |

> Cada captura muestra el panel de **Executions** de n8n con el estado de la corrida
> (success / error) y el detalle del nodo relevante.
