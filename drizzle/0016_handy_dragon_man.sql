CREATE TABLE `otReceptionHistory` (
	`id` int AUTO_INCREMENT NOT NULL,
	`workOrderId` int NOT NULL,
	`weekId` int NOT NULL,
	`previousStatus` enum('RECIBIDO','NO_RECIBIDO') NOT NULL,
	`newStatus` enum('RECIBIDO','NO_RECIBIDO') NOT NULL,
	`previousObservation` text,
	`newObservation` text,
	`facilitatorId` int NOT NULL,
	`facilitatorName` varchar(160) NOT NULL,
	`guard` enum('A','B','C','D') NOT NULL,
	`createdAt` timestamp NOT NULL DEFAULT (now()),
	CONSTRAINT `otReceptionHistory_id` PRIMARY KEY(`id`)
);
--> statement-breakpoint
CREATE TABLE `otReceptions` (
	`id` int AUTO_INCREMENT NOT NULL,
	`workOrderId` int NOT NULL,
	`weekId` int NOT NULL,
	`status` enum('RECIBIDO','NO_RECIBIDO') NOT NULL DEFAULT 'NO_RECIBIDO',
	`observation` text,
	`version` int NOT NULL DEFAULT 1,
	`updatedAt` timestamp NOT NULL DEFAULT (now()) ON UPDATE CURRENT_TIMESTAMP,
	`updatedById` int NOT NULL,
	`updatedByName` varchar(160) NOT NULL,
	`updatedGuard` enum('A','B','C','D') NOT NULL,
	CONSTRAINT `otReceptions_id` PRIMARY KEY(`id`),
	CONSTRAINT `otReceptions_workOrder_week_unique` UNIQUE(`workOrderId`,`weekId`)
);
--> statement-breakpoint
CREATE INDEX `otReceptionHistory_ot_week_created_idx` ON `otReceptionHistory` (`workOrderId`,`weekId`,`createdAt`);