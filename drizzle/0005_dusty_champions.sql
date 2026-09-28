CREATE TABLE `accessCodes` (
	`id` int AUTO_INCREMENT NOT NULL,
	`label` varchar(120) NOT NULL,
	`role` enum('facilitator','supervisor') NOT NULL,
	`salt` varchar(64) NOT NULL,
	`codeHash` varchar(128) NOT NULL,
	`active` boolean NOT NULL DEFAULT true,
	`createdById` int,
	`createdAt` timestamp NOT NULL DEFAULT (now()),
	`lastUsedAt` timestamp,
	CONSTRAINT `accessCodes_id` PRIMARY KEY(`id`),
	CONSTRAINT `accessCodes_hash_unique` UNIQUE(`codeHash`)
);
--> statement-breakpoint
CREATE INDEX `accessCodes_active_role_idx` ON `accessCodes` (`active`,`role`);