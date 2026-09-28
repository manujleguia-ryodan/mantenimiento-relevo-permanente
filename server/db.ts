import { and, desc, eq, inArray } from "drizzle-orm";
import { drizzle } from "drizzle-orm/mysql2";
import {
  InsertUser, equipment, guardAssignments, guardRelations, handoverItems, handovers,
  maintenanceWeeks, otHistory, otStatePendingWarehouses, otStates, users, warehouses, workOrders,
} from "../drizzle/schema";
import { ENV } from "./_core/env";
import { hasReservation } from "../shared/maintenance";

let _db: ReturnType<typeof drizzle> | null = null;
export async function getDb() {
  if (!_db && process.env.DATABASE_URL) {
    try { _db = drizzle(process.env.DATABASE_URL); } catch (error) { console.warn("[Database] Failed to connect:", error); _db = null; }
  }
  return _db;
}
export async function requireDb() { const db = await getDb(); if (!db) throw new Error("Database is not available"); return db; }

export async function upsertUser(user: InsertUser): Promise<void> {
  if (!user.openId) throw new Error("User openId is required for upsert");
  const db = await getDb(); if (!db) return;
  const values: InsertUser = { openId: user.openId }; const updateSet: Record<string, unknown> = {};
  for (const field of ["name", "email", "loginMethod"] as const) if (user[field] !== undefined) { values[field] = user[field] ?? null; updateSet[field] = user[field] ?? null; }
  values.lastSignedIn = user.lastSignedIn ?? new Date(); updateSet.lastSignedIn = values.lastSignedIn;
  if (user.role !== undefined) { values.role = user.role; updateSet.role = user.role; }
  else if (user.openId === ENV.ownerOpenId) { values.role = "admin"; updateSet.role = "admin"; }
  await db.insert(users).values(values).onDuplicateKeyUpdate({ set: updateSet });
}
export async function getUserByOpenId(openId: string) { const db = await getDb(); if (!db) return undefined; const [result] = await db.select().from(users).where(eq(users.openId, openId)).limit(1); return result; }
export async function getWeeks() { const db = await requireDb(); return db.select().from(maintenanceWeeks).orderBy(desc(maintenanceWeeks.importedAt)); }
export async function getEquipment() { const db = await requireDb(); return db.select().from(equipment).orderBy(equipment.code); }
export async function getWarehouses() { const db = await requireDb(); return db.select().from(warehouses).where(eq(warehouses.active, true)).orderBy(warehouses.name); }
export async function getGuardRelations() { const db = await requireDb(); return db.select().from(guardRelations).orderBy(guardRelations.fromGuard); }

export async function getDashboardData(weekId: number, guard: "A" | "B" | "C" | "D") {
  const db = await requireDb();
  const [week] = await db.select().from(maintenanceWeeks).where(eq(maintenanceWeeks.id, weekId)).limit(1);
  if (!week) throw new Error("Semana no encontrada");
  const selected = await db.select().from(guardAssignments).where(and(eq(guardAssignments.weekId, weekId), eq(guardAssignments.guard, guard)));
  const selectedIds = selected
    .map(item => item.equipmentId)
    .filter((id): id is number => Number.isSafeInteger(id) && id > 0);
  const rows = await db.select({ workOrder: workOrders, equipment, state: otStates, inheritedFrom: handovers.fromGuard })
    .from(workOrders).leftJoin(equipment, eq(workOrders.equipmentId, equipment.id))
    .leftJoin(otStates, and(eq(otStates.workOrderId, workOrders.id), eq(otStates.weekId, weekId), eq(otStates.guard, guard)))
    .leftJoin(handovers, eq(otStates.inheritedFromHandoverId, handovers.id)).where(eq(workOrders.weekId, weekId));
  const visibleRows = selectedIds.length ? rows.filter(row => selectedIds.includes(row.workOrder.equipmentId)) : [];
  const stateIds = visibleRows.map(row => row.state?.id).filter((id): id is number => Boolean(id));
  const pending = stateIds.length ? await db.select().from(otStatePendingWarehouses).where(inArray(otStatePendingWarehouses.otStateId, stateIds)) : [];
  const pendingByState = new Map<number, typeof pending>();
  for (const item of pending) { const list = pendingByState.get(item.otStateId) ?? []; list.push(item); pendingByState.set(item.otStateId, list); }
  const enrichedRows = visibleRows.map(row => ({
    ...row,
    requiresLogistics: hasReservation(row.workOrder.sourceData),
    pendingWarehouses: row.state?.id ? pendingByState.get(row.state.id) ?? [] : [],
  }));
  return { week, guard, selectedEquipmentIds: selectedIds, rows: enrichedRows, selectedEquipment: (await getEquipment()).filter(item => selectedIds.includes(item.id)), allEquipment: await getEquipment(), warehouses: await getWarehouses() };
}

export async function getOtHistory(workOrderId: number) { const db = await requireDb(); return db.select().from(otHistory).where(eq(otHistory.workOrderId, workOrderId)).orderBy(desc(otHistory.createdAt)); }
export async function getLatestHandover(weekId: number, fromGuard: "A" | "B" | "C" | "D", toGuard: "A" | "B" | "C" | "D") {
  const db = await requireDb(); const [handover] = await db.select().from(handovers).where(and(eq(handovers.weekId, weekId), eq(handovers.fromGuard, fromGuard), eq(handovers.toGuard, toGuard))).orderBy(desc(handovers.createdAt)).limit(1);
  if (!handover) return null; const items = await db.select().from(handoverItems).where(eq(handoverItems.handoverId, handover.id)); return { handover, items };
}
export async function getWorkOrdersByIds(ids: number[]) {
  const db = await requireDb(); if (!ids.length) return [];
  return db.select({ workOrder: workOrders, equipment, state: otStates }).from(workOrders).leftJoin(equipment, eq(workOrders.equipmentId, equipment.id)).leftJoin(otStates, eq(otStates.workOrderId, workOrders.id)).where(inArray(workOrders.id, ids));
}
