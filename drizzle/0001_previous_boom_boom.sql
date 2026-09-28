CREATE TABLE `equipment` (
	`id` int AUTO_INCREMENT NOT NULL,
	`code` varchar(80) NOT NULL,
	`description` varchar(255),
	`createdAt` timestamp NOT NULL DEFAULT (now()),
	CONSTRAINT `equipment_id` PRIMARY KEY(`id`),
	CONSTRAINT `equipment_code_unique` UNIQUE(`code`)
);
--> statement-breakpoint
CREATE TABLE `guardAssignments` (
	`id` int AUTO_INCREMENT NOT NULL,
	`weekId` int NOT NULL,
	`guard` enum('A','B','C','D') NOT NULL,
	`equipmentId` int NOT NULL,
	`selectedById` int,
	`selectedAt` timestamp NOT NULL DEFAULT (now()),
	CONSTRAINT `guardAssignments_id` PRIMARY KEY(`id`),
	CONSTRAINT `guardAssignments_unique` UNIQUE(`weekId`,`guard`,`equipmentId`)
);
--> statement-breakpoint
CREATE TABLE `guardRelations` (
	`id` int AUTO_INCREMENT NOT NULL,
	`fromGuard` enum('A','B','C','D') NOT NULL,
	`toGuard` enum('A','B','C','D') NOT NULL,
	`updatedById` int,
	`updatedAt` timestamp NOT NULL DEFAULT (now()) ON UPDATE CURRENT_TIMESTAMP,
	CONSTRAINT `guardRelations_id` PRIMARY KEY(`id`),
	CONSTRAINT `guardRelations_from_unique` UNIQUE(`fromGuard`)
);
--> statement-breakpoint
CREATE TABLE `handoverItems` (
	`id` int AUTO_INCREMENT NOT NULL,
	`handoverId` int NOT NULL,
	`otStateId` int NOT NULL,
	CONSTRAINT `handoverItems_id` PRIMARY KEY(`id`),
	CONSTRAINT `handoverItems_unique` UNIQUE(`handoverId`,`otStateId`)
);
--> statement-breakpoint
CREATE TABLE `handovers` (
	`id` int AUTO_INCREMENT NOT NULL,
	`weekId` int NOT NULL,
	`fromGuard` enum('A','B','C','D') NOT NULL,
	`toGuard` enum('A','B','C','D') NOT NULL,
	`status` enum('BORRADOR','ENVIADO','CONFIRMADO','REABIERTO','ANULADO') NOT NULL DEFAULT 'BORRADOR',
	`createdById` int,
	`createdAt` timestamp NOT NULL DEFAULT (now()),
	`confirmedById` int,
	`confirmedAt` timestamp,
	CONSTRAINT `handovers_id` PRIMARY KEY(`id`)
);
--> statement-breakpoint
CREATE TABLE `maintenanceWeeks` (
	`id` int AUTO_INCREMENT NOT NULL,
	`label` varchar(32) NOT NULL,
	`weekNumber` int NOT NULL,
	`year` int NOT NULL,
	`sourceFileName` varchar(255),
	`importedById` int,
	`importedAt` timestamp NOT NULL DEFAULT (now()),
	`isArchived` boolean NOT NULL DEFAULT false,
	CONSTRAINT `maintenanceWeeks_id` PRIMARY KEY(`id`),
	CONSTRAINT `maintenanceWeeks_period_unique` UNIQUE(`year`,`weekNumber`)
);
--> statement-breakpoint
CREATE TABLE `otHistory` (
	`id` int AUTO_INCREMENT NOT NULL,
	`otStateId` int NOT NULL,
	`workOrderId` int NOT NULL,
	`weekId` int NOT NULL,
	`guard` enum('A','B','C','D') NOT NULL,
	`previousStatus` varchar(32),
	`newStatus` varchar(32) NOT NULL,
	`previousPendingWarehouse` varchar(32),
	`newPendingWarehouse` varchar(32),
	`note` text,
	`actorUserId` int,
	`actorName` varchar(160),
	`createdAt` timestamp NOT NULL DEFAULT (now()),
	CONSTRAINT `otHistory_id` PRIMARY KEY(`id`)
);
--> statement-breakpoint
CREATE TABLE `otStates` (
	`id` int AUTO_INCREMENT NOT NULL,
	`workOrderId` int NOT NULL,
	`weekId` int NOT NULL,
	`guard` enum('A','B','C','D') NOT NULL,
	`status` enum('SIN_REVISAR','RECIBIDA','PARCIAL','FALTA_RECIBIR') NOT NULL DEFAULT 'SIN_REVISAR',
	`pendingWarehouse1` boolean NOT NULL DEFAULT false,
	`pendingWarehouse2` boolean NOT NULL DEFAULT false,
	`responsibleName` varchar(160),
	`note` text,
	`updatedById` int,
	`updatedAt` timestamp NOT NULL DEFAULT (now()) ON UPDATE CURRENT_TIMESTAMP,
	CONSTRAINT `otStates_id` PRIMARY KEY(`id`),
	CONSTRAINT `otStates_unique` UNIQUE(`workOrderId`,`weekId`,`guard`)
);
--> statement-breakpoint
CREATE TABLE `workOrders` (
	`id` int AUTO_INCREMENT NOT NULL,
	`weekId` int NOT NULL,
	`otNumber` varchar(80) NOT NULL,
	`equipmentId` int NOT NULL,
	`description` text,
	`sourceData` text,
	`createdAt` timestamp NOT NULL DEFAULT (now()),
	CONSTRAINT `workOrders_id` PRIMARY KEY(`id`),
	CONSTRAINT `workOrders_period_ot_unique` UNIQUE(`weekId`,`otNumber`)
);
--> statement-breakpoint
CREATE INDEX `guardAssignments_lookup_idx` ON `guardAssignments` (`weekId`,`guard`);--> statement-breakpoint
CREATE INDEX `handovers_week_idx` ON `handovers` (`weekId`,`fromGuard`,`toGuard`);--> statement-breakpoint
CREATE INDEX `maintenanceWeeks_period_idx` ON `maintenanceWeeks` (`year`,`weekNumber`);--> statement-breakpoint
CREATE INDEX `otHistory_ot_idx` ON `otHistory` (`workOrderId`,`createdAt`);--> statement-breakpoint
CREATE INDEX `otHistory_week_idx` ON `otHistory` (`weekId`,`guard`);--> statement-breakpoint
CREATE INDEX `otStates_lookup_idx` ON `otStates` (`weekId`,`guard`,`status`);--> statement-breakpoint
CREATE INDEX `workOrders_week_idx` ON `workOrders` (`weekId`);--> statement-breakpoint
CREATE INDEX `workOrders_equipment_idx` ON `workOrders` (`equipmentId`);