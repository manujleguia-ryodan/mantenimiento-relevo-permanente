CREATE TABLE `otExecutions` (
	`id` int AUTO_INCREMENT NOT NULL,
	`workOrderId` int NOT NULL,
	`weekId` int NOT NULL,
	`guard` enum('A','B','C','D') NOT NULL,
	`status` enum('ABIERTA','CERRADA_NOTIFICADA') NOT NULL DEFAULT 'ABIERTA',
	`executedById` int,
	`executedByName` varchar(160),
	`executedAt` timestamp NOT NULL DEFAULT (now()) ON UPDATE CURRENT_TIMESTAMP,
	`note` text,
	CONSTRAINT `otExecutions_id` PRIMARY KEY(`id`),
	CONSTRAINT `otExecutions_unique` UNIQUE(`workOrderId`,`weekId`,`guard`)
);
--> statement-breakpoint
CREATE INDEX `otExecutions_lookup_idx` ON `otExecutions` (`weekId`,`guard`,`status`);