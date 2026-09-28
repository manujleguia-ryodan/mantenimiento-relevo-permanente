CREATE TABLE `otStatePendingWarehouses` (
	`id` int AUTO_INCREMENT NOT NULL,
	`otStateId` int NOT NULL,
	`warehouseId` int NOT NULL,
	`warehouseCodeSnapshot` varchar(64) NOT NULL,
	`warehouseNameSnapshot` varchar(160) NOT NULL,
	`createdAt` timestamp NOT NULL DEFAULT (now()),
	CONSTRAINT `otStatePendingWarehouses_id` PRIMARY KEY(`id`),
	CONSTRAINT `otStatePendingWarehouses_unique` UNIQUE(`otStateId`,`warehouseId`)
);
--> statement-breakpoint
CREATE TABLE `warehouses` (
	`id` int AUTO_INCREMENT NOT NULL,
	`code` varchar(64) NOT NULL,
	`name` varchar(160) NOT NULL,
	`active` boolean NOT NULL DEFAULT true,
	`createdAt` timestamp NOT NULL DEFAULT (now()),
	CONSTRAINT `warehouses_id` PRIMARY KEY(`id`),
	CONSTRAINT `warehouses_code_unique` UNIQUE(`code`)
);
--> statement-breakpoint
ALTER TABLE `handoverItems` MODIFY COLUMN `otStateId` int;--> statement-breakpoint
ALTER TABLE `otHistory` MODIFY COLUMN `previousPendingWarehouse` varchar(200);--> statement-breakpoint
ALTER TABLE `otHistory` MODIFY COLUMN `newPendingWarehouse` varchar(200);--> statement-breakpoint
ALTER TABLE `handoverItems` ADD `workOrderId` int;--> statement-breakpoint
ALTER TABLE `handoverItems` ADD `statusSnapshot` varchar(32);--> statement-breakpoint
ALTER TABLE `handoverItems` ADD `pendingWarehousesSnapshot` text;--> statement-breakpoint
ALTER TABLE `handoverItems` ADD `fromGuardSnapshot` enum('A','B','C','D');--> statement-breakpoint
ALTER TABLE `handoverItems` ADD `toGuardSnapshot` enum('A','B','C','D');--> statement-breakpoint
ALTER TABLE `handoverItems` ADD `createdByIdSnapshot` int;--> statement-breakpoint
ALTER TABLE `handoverItems` ADD `createdAtSnapshot` timestamp;--> statement-breakpoint
ALTER TABLE `otStates` ADD `inheritedFromHandoverId` int;--> statement-breakpoint
CREATE INDEX `otStatePendingWarehouses_state_idx` ON `otStatePendingWarehouses` (`otStateId`);--> statement-breakpoint
CREATE INDEX `otStates_inherited_idx` ON `otStates` (`inheritedFromHandoverId`);