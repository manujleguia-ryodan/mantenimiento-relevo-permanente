import {
  boolean,
  index,
  int,
  mysqlEnum,
  mysqlTable,
  text,
  timestamp,
  uniqueIndex,
  varchar,
} from "drizzle-orm/mysql-core";

export const users = mysqlTable("users", {
  id: int("id").autoincrement().primaryKey(),
  openId: varchar("openId", { length: 64 }).notNull().unique(),
  name: text("name"), email: varchar("email", { length: 320 }),
  loginMethod: varchar("loginMethod", { length: 64 }),
  role: mysqlEnum("role", ["user", "admin", "facilitator", "supervisor", "viewer"]).default("user").notNull(),
  createdAt: timestamp("createdAt").defaultNow().notNull(),
  updatedAt: timestamp("updatedAt").defaultNow().onUpdateNow().notNull(),
  lastSignedIn: timestamp("lastSignedIn").defaultNow().notNull(),
});

export const accessCodes = mysqlTable("accessCodes", {
  id: int("id").autoincrement().primaryKey(),
  label: varchar("label", { length: 120 }).notNull(),
  role: mysqlEnum("role", ["facilitator", "supervisor"]).notNull(),
  guard: mysqlEnum("guard", ["A", "B", "C", "D"]).default("A").notNull(),
  salt: varchar("salt", { length: 64 }).notNull(),
  codeHash: varchar("codeHash", { length: 128 }).notNull(),
  active: boolean("active").default(true).notNull(),
  createdById: int("createdById"),
  createdAt: timestamp("createdAt").defaultNow().notNull(),
  lastUsedAt: timestamp("lastUsedAt"),
}, table => ({
  codeHashUnique: uniqueIndex("accessCodes_hash_unique").on(table.codeHash),
  activeRoleIdx: index("accessCodes_active_role_idx").on(table.active, table.role),
}));

export const maintenanceWeeks = mysqlTable("maintenanceWeeks", {
  id: int("id").autoincrement().primaryKey(), label: varchar("label", { length: 32 }).notNull(),
  weekNumber: int("weekNumber").notNull(), year: int("year").notNull(),
  sourceFileName: varchar("sourceFileName", { length: 255 }), importedById: int("importedById"),
  importedAt: timestamp("importedAt").defaultNow().notNull(), isArchived: boolean("isArchived").default(false).notNull(),
}, table => ({
  periodUnique: uniqueIndex("maintenanceWeeks_period_unique").on(table.year, table.weekNumber),
  periodIdx: index("maintenanceWeeks_period_idx").on(table.year, table.weekNumber),
}));

export const equipment = mysqlTable("equipment", {
  id: int("id").autoincrement().primaryKey(), code: varchar("code", { length: 80 }).notNull(), description: varchar("description", { length: 255 }), activo_en_taller: boolean("activo_en_taller").default(false).notNull(), createdAt: timestamp("createdAt").defaultNow().notNull(),
}, table => ({ codeUnique: uniqueIndex("equipment_code_unique").on(table.code) }));

export const warehouses = mysqlTable("warehouses", {
  id: int("id").autoincrement().primaryKey(), code: varchar("code", { length: 64 }).notNull(),
  name: varchar("name", { length: 160 }).notNull(), active: boolean("active").default(true).notNull(),
  createdAt: timestamp("createdAt").defaultNow().notNull(),
}, table => ({ codeUnique: uniqueIndex("warehouses_code_unique").on(table.code) }));

export const workOrders = mysqlTable("workOrders", {
  id: int("id").autoincrement().primaryKey(), weekId: int("weekId").notNull(),
  otNumber: varchar("otNumber", { length: 80 }).notNull(), equipmentId: int("equipmentId").notNull(),
  description: text("description"), sourceData: text("sourceData"),
  isCancelled: boolean("isCancelled").default(false).notNull(),
  cancelledAt: timestamp("cancelledAt"), cancelledById: int("cancelledById"),
  cancellationReason: varchar("cancellationReason", { length: 500 }),
  createdAt: timestamp("createdAt").defaultNow().notNull(),
}, table => ({
  otNumberUnique: uniqueIndex("workOrders_ot_number_unique").on(table.otNumber),
  weekIdx: index("workOrders_week_idx").on(table.weekId), equipmentIdx: index("workOrders_equipment_idx").on(table.equipmentId),
}));

export const otExecutions = mysqlTable("otExecutions", {
  id: int("id").autoincrement().primaryKey(), workOrderId: int("workOrderId").notNull(), weekId: int("weekId").notNull(),
  guard: mysqlEnum("guard", ["A", "B", "C", "D"]).notNull(), status: mysqlEnum("status", ["ABIERTA", "CERRADA_NOTIFICADA"]).default("ABIERTA").notNull(),
  executedById: int("executedById"), executedByName: varchar("executedByName", { length: 160 }), executedAt: timestamp("executedAt").defaultNow().onUpdateNow().notNull(), note: text("note"),
}, table => ({
  executionUnique: uniqueIndex("otExecutions_unique").on(table.workOrderId, table.weekId, table.guard),
  executionLookupIdx: index("otExecutions_lookup_idx").on(table.weekId, table.guard, table.status),
}));

export const guardAssignments = mysqlTable("guardAssignments", {
  id: int("id").autoincrement().primaryKey(), weekId: int("weekId").notNull(),
  guard: mysqlEnum("guard", ["A", "B", "C", "D"]).notNull(), equipmentId: int("equipmentId").notNull(),
  selectedById: int("selectedById"), selectedAt: timestamp("selectedAt").defaultNow().notNull(),
}, table => ({
  assignmentUnique: uniqueIndex("guardAssignments_unique").on(table.weekId, table.guard, table.equipmentId),
  lookupIdx: index("guardAssignments_lookup_idx").on(table.weekId, table.guard),
}));

/** Selección privada de equipos de cada Supervisor; no comparte asignaciones del Facilitador. */
export const supervisorAssignments = mysqlTable("supervisorAssignments", {
  id: int("id").autoincrement().primaryKey(),
  supervisorName: varchar("supervisorName", { length: 160 }).notNull(),
  weekId: int("weekId").notNull(),
  guard: mysqlEnum("guard", ["A", "B", "C", "D"]).notNull(),
  equipmentId: int("equipmentId").notNull(),
  selectedAt: timestamp("selectedAt").defaultNow().notNull(),
}, table => ({
  assignmentUnique: uniqueIndex("supervisorAssignments_unique").on(table.supervisorName, table.weekId, table.guard, table.equipmentId),
  lookupIdx: index("supervisorAssignments_lookup_idx").on(table.supervisorName, table.weekId, table.guard),
}));

export const otStates = mysqlTable("otStates", {
  id: int("id").autoincrement().primaryKey(), workOrderId: int("workOrderId").notNull(), weekId: int("weekId").notNull(),
  guard: mysqlEnum("guard", ["A", "B", "C", "D"]).notNull(),
  status: mysqlEnum("status", ["SIN_REVISAR", "RECIBIDA", "PARCIAL", "FALTA_RECIBIR"]).default("SIN_REVISAR").notNull(),
  pendingWarehouse1: boolean("pendingWarehouse1").default(false).notNull(), pendingWarehouse2: boolean("pendingWarehouse2").default(false).notNull(),
  inheritedFromHandoverId: int("inheritedFromHandoverId"), responsibleName: varchar("responsibleName", { length: 160 }), reviewedByRole: varchar("reviewedByRole", { length: 32 }),
  note: text("note"), updatedById: int("updatedById"), updatedAt: timestamp("updatedAt").defaultNow().onUpdateNow().notNull(),
}, table => ({
  stateUnique: uniqueIndex("otStates_unique").on(table.workOrderId, table.weekId, table.guard),
  stateLookupIdx: index("otStates_lookup_idx").on(table.weekId, table.guard, table.status),
  inheritedIdx: index("otStates_inherited_idx").on(table.inheritedFromHandoverId),
}));

export const otStatePendingWarehouses = mysqlTable("otStatePendingWarehouses", {
  id: int("id").autoincrement().primaryKey(), otStateId: int("otStateId").notNull(), warehouseId: int("warehouseId").notNull(),
  warehouseCodeSnapshot: varchar("warehouseCodeSnapshot", { length: 64 }).notNull(), warehouseNameSnapshot: varchar("warehouseNameSnapshot", { length: 160 }).notNull(),
  createdAt: timestamp("createdAt").defaultNow().notNull(),
}, table => ({
  itemUnique: uniqueIndex("otStatePendingWarehouses_unique").on(table.otStateId, table.warehouseId),
  stateIdx: index("otStatePendingWarehouses_state_idx").on(table.otStateId),
}));

export const otHistory = mysqlTable("otHistory", {
  id: int("id").autoincrement().primaryKey(), otStateId: int("otStateId").notNull(), workOrderId: int("workOrderId").notNull(),
  weekId: int("weekId").notNull(), guard: mysqlEnum("guard", ["A", "B", "C", "D"]).notNull(),
  previousStatus: varchar("previousStatus", { length: 32 }), newStatus: varchar("newStatus", { length: 32 }).notNull(),
  previousPendingWarehouse: varchar("previousPendingWarehouse", { length: 200 }), newPendingWarehouse: varchar("newPendingWarehouse", { length: 200 }),
  note: text("note"), actorUserId: int("actorUserId"), actorName: varchar("actorName", { length: 160 }), createdAt: timestamp("createdAt").defaultNow().notNull(),
}, table => ({
  historyOtIdx: index("otHistory_ot_idx").on(table.workOrderId, table.createdAt), historyWeekIdx: index("otHistory_week_idx").on(table.weekId, table.guard),
}));

export const guardRelations = mysqlTable("guardRelations", {
  id: int("id").autoincrement().primaryKey(), fromGuard: mysqlEnum("fromGuard", ["A", "B", "C", "D"]).notNull(),
  toGuard: mysqlEnum("toGuard", ["A", "B", "C", "D"]).notNull(), updatedById: int("updatedById"), updatedAt: timestamp("updatedAt").defaultNow().onUpdateNow().notNull(),
}, table => ({ fromGuardUnique: uniqueIndex("guardRelations_from_unique").on(table.fromGuard) }));

export const handovers = mysqlTable("handovers", {
  id: int("id").autoincrement().primaryKey(), weekId: int("weekId").notNull(),
  fromGuard: mysqlEnum("fromGuard", ["A", "B", "C", "D"]).notNull(), toGuard: mysqlEnum("toGuard", ["A", "B", "C", "D"]).notNull(),
  status: mysqlEnum("status", ["BORRADOR", "ENVIADO", "CONFIRMADO", "REABIERTO", "ANULADO"]).default("BORRADOR").notNull(),
  createdById: int("createdById"), createdAt: timestamp("createdAt").defaultNow().notNull(), confirmedById: int("confirmedById"), confirmedAt: timestamp("confirmedAt"),
}, table => ({ handoverIdx: index("handovers_week_idx").on(table.weekId, table.fromGuard, table.toGuard) }));

export const handoverItems = mysqlTable("handoverItems", {
  id: int("id").autoincrement().primaryKey(), handoverId: int("handoverId").notNull(), otStateId: int("otStateId"), workOrderId: int("workOrderId"),
  statusSnapshot: varchar("statusSnapshot", { length: 32 }), pendingWarehousesSnapshot: text("pendingWarehousesSnapshot"),
  fromGuardSnapshot: mysqlEnum("fromGuardSnapshot", ["A", "B", "C", "D"]), toGuardSnapshot: mysqlEnum("toGuardSnapshot", ["A", "B", "C", "D"]),
  createdByIdSnapshot: int("createdByIdSnapshot"), createdAtSnapshot: timestamp("createdAtSnapshot"),
}, table => ({ itemUnique: uniqueIndex("handoverItems_unique").on(table.handoverId, table.otStateId) }));

export type User = typeof users.$inferSelect;
export type InsertUser = typeof users.$inferInsert;
export type MaintenanceWeek = typeof maintenanceWeeks.$inferSelect;
export type Equipment = typeof equipment.$inferSelect;
export type Warehouse = typeof warehouses.$inferSelect;
export type WorkOrder = typeof workOrders.$inferSelect;
export type OtExecution = typeof otExecutions.$inferSelect;
export type OtState = typeof otStates.$inferSelect;
