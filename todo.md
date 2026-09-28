# Seguimiento de implementación

## Completado

- [x] Importación de Excel con equipos normalizados desde **DESCRIPCIÓN DE UBICACIÓN TÉCNICA** (`EQ_OCR02 → CR02`, `EQ_OCA13 → CA13`).
- [x] Tabla operacional limitada a OT, Equipo, Reserva y Descripción de operación, con respaldo para columnas ausentes.
- [x] Verificación de datos reales: las **222 OTs importadas** preservan el encabezado `RESERVA`; las filas vacías se excluyen de las métricas logísticas.
- [x] Regla logística: las OTs sin Reserva no exponen recepción/almacenes, no entran en KPIs de recepción y no se heredan en relevos.
- [x] Roles funcionales: Administrador, Facilitador, Supervisor y Visualizador, protegidos en interfaz y backend.
- [x] Acción independiente para Supervisor: **Cerrar y notificar OT**, sin modificar la recepción logística; Visualizador no puede accionarla.
- [x] Gestión de relaciones de guardia y accesos disponible solo para Administrador.
- [x] Validaciones automáticas de TypeScript, pruebas y build.

## Configuración posterior al desarrollo

El Administrador puede asignar roles desde **Gestionar accesos** cuando cada persona haya iniciado sesión con su propia cuenta. Esta asignación de personas reales depende de los usuarios que la organización incorpore, no de una función pendiente del sistema.
