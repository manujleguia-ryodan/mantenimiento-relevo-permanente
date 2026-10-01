CREATE TABLE `supervisorAssignments` (
	`id` int AUTO_INCREMENT NOT NULL,
	`supervisorName` varchar(160) NOT NULL,
	`weekId` int NOT NULL,
	`guard` enum('A','B','C','D') NOT NULL,
	`equipmentId` int NOT NULL,
	`selectedAt` timestamp NOT NULL DEFAULT (now()),
	CONSTRAINT `supervisorAssignments_id` PRIMARY KEY(`id`),
	CONSTRAINT `supervisorAssignments_unique` UNIQUE(`supervisorName`,`weekId`,`guard`,`equipmentId`)
);
--> statement-breakpoint
ALTER TABLE `equipment` ADD `activo_en_taller` boolean DEFAULT false NOT NULL;--> statement-breakpoint
CREATE INDEX `supervisorAssignments_lookup_idx` ON `supervisorAssignments` (`supervisorName`,`weekId`,`guard`);