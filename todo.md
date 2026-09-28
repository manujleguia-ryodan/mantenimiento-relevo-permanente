# Seguimiento de implementación

## Completado
- [x] Importación de Excel con equipos normalizados desde **DESCRIPCIÓN DE UBICACIÓN TÉCNICA** (`EQ_OCR02 → CR02`, `EQ_OCA13 → CA13`).
- [x] Tabla operacional limitada a OT, Equipo, Reserva y Descripción de operación, con respaldo para columnas ausentes.
- [x] Reglas de Reserva: las OTs sin Reserva se excluyen de estados, almacenes, KPIs logísticos y relevo.
- [x] Roles funcionales: Administrador, Facilitador, Supervisor y Visualizador, protegidos en interfaz y backend.
- [x] Administración de relaciones de guardia y usuarios disponible solo para Administrador.
- [x] Validaciones automáticas de TypeScript, pruebas y build.

## Pendiente de configuración operativa
- [ ] Iniciar sesión con cuentas Manus de los usuarios y asignar sus roles desde **Gestionar accesos** (Administrador).
- [ ] Verificar en operación que el archivo semanal entregue el encabezado Reserva cuando corresponda.
- [ ] Definir, si se necesita, una acción de campo específica para Supervisor (por ejemplo, marcar ejecución física), independiente de la recepción logística.
