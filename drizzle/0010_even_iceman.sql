ALTER TABLE `workOrders` DROP INDEX `workOrders_period_ot_unique`;--> statement-breakpoint
ALTER TABLE `workOrders` ADD CONSTRAINT `workOrders_ot_number_unique` UNIQUE(`otNumber`);