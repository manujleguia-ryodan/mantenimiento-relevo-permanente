CREATE TABLE `guardWorkOrderAssignments` (
	`id` int AUTO_INCREMENT NOT NULL,
	`weekId` int NOT NULL,
	`guard` enum('A','B','C','D') NOT NULL,
	`workOrderId` int NOT NULL,
	`selectedById` int,
	`selectedAt` timestamp NOT NULL DEFAULT (now()),
	CONSTRAINT `guardWorkOrderAssignments_id` PRIMARY KEY(`id`),
	CONSTRAINT `guardWorkOrderAssignments_unique` UNIQUE(`weekId`,`guard`,`workOrderId`)
);
--> statement-breakpoint
CREATE INDEX `guardWorkOrderAssignments_lookup_idx` ON `guardWorkOrderAssignments` (`weekId`,`guard`);