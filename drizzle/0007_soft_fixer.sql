ALTER TABLE `workOrders` ADD `isCancelled` boolean DEFAULT false NOT NULL;--> statement-breakpoint
ALTER TABLE `workOrders` ADD `cancelledAt` timestamp;--> statement-breakpoint
ALTER TABLE `workOrders` ADD `cancelledById` int;--> statement-breakpoint
ALTER TABLE `workOrders` ADD `cancellationReason` varchar(500);