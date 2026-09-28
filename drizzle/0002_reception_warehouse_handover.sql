-- Fase 2: almacenes controlados, herencia y snapshots de relevo.
-- Esta migración no cambia workOrders.id ni elimina datos existentes.

CREATE TABLE `warehouses` (
  `id` int AUTO_INCREMENT NOT NULL,
  `code` varchar(64) NOT NULL,
  `name` varchar(160) NOT NULL,
  `active` boolean NOT NULL DEFAULT true,
  `createdAt` timestamp NOT NULL DEFAULT (now()),
  CONSTRAINT `warehouses_id` PRIMARY KEY(`id`),
  CONSTRAINT `warehouses_code_unique` UNIQUE(`code`)
);

INSERT IGNORE INTO `warehouses` (`code`, `name`, `active`) VALUES
  ('CENTRAL', 'Central', true),
  ('FERREYROS', 'Ferreyros', true),
  ('DONALDSON', 'Donaldson', true);

ALTER TABLE `otStates`
  ADD COLUMN `inheritedFromHandoverId` int NULL;

CREATE INDEX `otStates_inherited_idx` ON `otStates` (`inheritedFromHandoverId`);

CREATE TABLE `otStatePendingWarehouses` (
  `id` int AUTO_INCREMENT NOT NULL,
  `otStateId` int NOT NULL,
  `warehouseId` int NOT NULL,
  `warehouseCodeSnapshot` varchar(64) NOT NULL,
  `warehouseNameSnapshot` varchar(160) NOT NULL,
  `createdAt` timestamp NOT NULL DEFAULT (now()),
  CONSTRAINT `otStatePendingWarehouses_id` PRIMARY KEY(`id`),
  CONSTRAINT `otStatePendingWarehouses_unique` UNIQUE(`otStateId`, `warehouseId`)
);

CREATE INDEX `otStatePendingWarehouses_state_idx` ON `otStatePendingWarehouses` (`otStateId`);

ALTER TABLE `handoverItems`
  MODIFY COLUMN `otStateId` int NULL,
  ADD COLUMN `workOrderId` int NULL,
  ADD COLUMN `statusSnapshot` varchar(32) NULL,
  ADD COLUMN `pendingWarehousesSnapshot` text NULL,
  ADD COLUMN `fromGuardSnapshot` enum('A','B','C','D') NULL,
  ADD COLUMN `toGuardSnapshot` enum('A','B','C','D') NULL,
  ADD COLUMN `createdByIdSnapshot` int NULL,
  ADD COLUMN `createdAtSnapshot` timestamp NULL;

-- Los flags legacy se conservan para no perder información. Como no identifican
-- un almacén real, no se convierten automáticamente a Central/Ferreyros/Donaldson.
-- La aplicación los tratará como registros históricos Pendientes hasta revisión.

-- Rollback manual si fuera necesario:
-- DROP TABLE `otStatePendingWarehouses`;
-- ALTER TABLE `otStates` DROP INDEX `otStates_inherited_idx`, DROP COLUMN `inheritedFromHandoverId`;
-- ALTER TABLE `handoverItems` DROP COLUMN `workOrderId`, DROP COLUMN `statusSnapshot`,
--   DROP COLUMN `pendingWarehousesSnapshot`, DROP COLUMN `fromGuardSnapshot`,
--   DROP COLUMN `toGuardSnapshot`, DROP COLUMN `createdByIdSnapshot`, DROP COLUMN `createdAtSnapshot`;
-- DROP TABLE `warehouses`;
