-- SPRINT 1: recepción compartida por OT y semana.
-- El respaldo debe existir antes de ejecutar esta migración.
-- Se conserva el estado más reciente por updatedAt y, en empate, por id.

CREATE TEMPORARY TABLE `otState_shared_reception_map` (
  `loserId` int NOT NULL PRIMARY KEY,
  `winnerId` int NOT NULL
);--> statement-breakpoint

INSERT INTO `otState_shared_reception_map` (`loserId`, `winnerId`)
SELECT loser.`id`, winner.`id`
FROM `otStates` AS loser
INNER JOIN `otStates` AS winner
  ON winner.`workOrderId` = loser.`workOrderId`
 AND winner.`weekId` = loser.`weekId`
 AND winner.`id` <> loser.`id`
LEFT JOIN `otStates` AS newer
  ON newer.`workOrderId` = winner.`workOrderId`
 AND newer.`weekId` = winner.`weekId`
 AND (
      newer.`updatedAt` > winner.`updatedAt`
      OR (newer.`updatedAt` = winner.`updatedAt` AND newer.`id` > winner.`id`)
    )
WHERE newer.`id` IS NULL;--> statement-breakpoint

-- Reubicar almacenes pendientes al estado ganador sin duplicar warehouseId.
INSERT IGNORE INTO `otStatePendingWarehouses`
  (`otStateId`, `warehouseId`, `warehouseCodeSnapshot`, `warehouseNameSnapshot`, `createdAt`)
SELECT map.`winnerId`, pending.`warehouseId`, pending.`warehouseCodeSnapshot`, pending.`warehouseNameSnapshot`, pending.`createdAt`
FROM `otState_shared_reception_map` AS map
INNER JOIN `otStatePendingWarehouses` AS pending
  ON pending.`otStateId` = map.`loserId`;--> statement-breakpoint

-- Todos los eventos de auditoría se conservan y pasan a referenciar el estado ganador.
UPDATE `otHistory` AS history
INNER JOIN `otState_shared_reception_map` AS map
  ON map.`loserId` = history.`otStateId`
SET history.`otStateId` = map.`winnerId`;--> statement-breakpoint

DELETE pending
FROM `otStatePendingWarehouses` AS pending
INNER JOIN `otState_shared_reception_map` AS map
  ON map.`loserId` = pending.`otStateId`;--> statement-breakpoint

DELETE state
FROM `otStates` AS state
INNER JOIN `otState_shared_reception_map` AS map
  ON map.`loserId` = state.`id`;--> statement-breakpoint

DROP TEMPORARY TABLE `otState_shared_reception_map`;--> statement-breakpoint

ALTER TABLE `otStates` DROP INDEX `otStates_unique`;--> statement-breakpoint
ALTER TABLE `otStates` ADD CONSTRAINT `otStates_unique` UNIQUE(`workOrderId`,`weekId`);
