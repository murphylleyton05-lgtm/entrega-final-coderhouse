# Evidencias de ejecución

Esta carpeta contiene la evidencia visual del sistema.

| Archivo | Qué muestra |
|---------|-------------|
| `flujo_completo_n8n.jpeg` | El flujo completo armado y desplegado en n8n cloud, con todos los nodos: Webhook, Airtable, Claude, Slack (HITL), Gmail y el Error Handler |

## Test de estrés — escenarios probados

Los 5 escenarios del test de estrés (camino feliz e infeliz) están documentados en detalle
en el PDF de arquitectura, sección 6 (Gestión de Errores y Resiliencia):

1. Lead válido → clasificado y enviado (camino feliz)
2. Lead VIP con score alto → validación HITL
3. Lead estándar → archivado sin contacto
4. Email vacío → el filtro detiene el flujo (camino infeliz)
5. Fallo de API de Claude → se registra en `Error_log` (camino infeliz)

> Para reforzar la evidencia, se pueden agregar aquí las capturas del panel
> **Executions** de n8n de cada corrida.
