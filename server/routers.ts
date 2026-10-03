import { TRPCError } from "@trpc/server";
import { and, desc, eq, inArray } from "drizzle-orm";
import { z } from "zod";
import { COOKIE_NAME } from "@shared/const";
import {
  accessCodes,
  equipment,
  guardAssignments,
  guardRelations,
  handoverItems,
  handovers,
  maintenanceWeeks,
  otExecutions,
  otHistory,
  otStatePendingWarehouses,
  otStates,
  supervisorAssignments,
  users,
  warehouses,
  workOrders,
} from "../drizzle/schema";
import {
  getDashboardData,
  getDb,
  getEquipment,
  getGuardRelations,
  getLatestHandover,
  getOtHistory,
  searchHistoricalWorkOrders,
  getWarehouses,
  getWeeks,
  requireDb,
} from "./db";
import { getSessionCookieOptions } from "./_core/cookies";
import { systemRouter } from "./_core/systemRouter";
import { publicProcedure, router } from "./_core/trpc";
import { ENV } from "./_core/env";
import {
  areAllWorkOrdersClosed,
  canPurgeSapRescheduledWorkOrder,
  canReviewReception,
  hasReservation,
  validateReceptionInput,
} from "@shared/maintenance";
import {
  canAttemptSharedPin,
  clearFailedSharedPinAttempts,
  createSharedAccessToken,
  createSharedAccessUser,
  getConfiguredSharedPin,
  getSharedAccessGuard,
  getSharedAccessName,
  getSharedAccessRole,
  hashAccessCode,
  hasSharedAccess,
  matchesAccessCode,
  matchesSharedPin,
  recordFailedSharedPinAttempt,
  SHARED_ACCESS_COOKIE,
  SHARED_ACCESS_TTL_MS,
  SHARED_PIN_MIN_LENGTH,
} from "./sharedAccess";

const guardSchema = z.enum(["A", "B", "C", "D"]);
const statusSchema = z.enum([
  "SIN_REVISAR",
  "RECIBIDA",
  "PARCIAL",
  "FALTA_RECIBIR",
]);
const managedRoleSchema = z.enum(["admin", "facilitator", "supervisor", "viewer"]);

const sharedAccessProcedure = publicProcedure;
const requireSession = publicProcedure.use(({ ctx, next }) => {
  if (!ctx.user) throw new TRPCError({ code: "UNAUTHORIZED", message: "Ingresa tu PIN para continuar." });
  return next();
});
const adminProcedure = requireSession.use(({ ctx, next }) => {
  if (ctx.user!.role !== "admin") throw new TRPCError({ code: "FORBIDDEN", message: "Solo el PIN maestro puede gestionar usuarios." });
  return next();
});
const facilitatorProcedure = requireSession.use(({ ctx, next }) => {
  if (ctx.user!.role !== "admin" && ctx.user!.role !== "facilitator") throw new TRPCError({ code: "FORBIDDEN", message: "Esta acción requiere una sesión de Facilitador." });
  return next();
});
const operationalProcedure = facilitatorProcedure;
const supervisorProcedure = facilitatorProcedure;

export function insertId(result: unknown) {
  const header = Array.isArray(result) ? result[0] : result;
  const value = (header as { insertId?: number | string | bigint } | null | undefined)?.insertId;
  const id = Number(value ?? 0);
  if (!Number.isSafeInteger(id) || id <= 0) {
    throw new Error("La base de datos no devolvió un ID de inserción válido.");
  }
  return id;
}

function correctiveSource(sourceData: string | null) {
  if (!sourceData) return null;
  try {
    const parsed = JSON.parse(sourceData) as Record<string, unknown>;
    return parsed.Tipo === "Correctiva" ? parsed : null;
  } catch {
    return null;
  }
}

function parseSourceData(sourceData: string | null): Record<string, string> {
  if (!sourceData) return {};
  try {
    const parsed = JSON.parse(sourceData) as unknown;
    return parsed && typeof parsed === "object" && !Array.isArray(parsed)
      ? Object.fromEntries(Object.entries(parsed as Record<string, unknown>).map(([key, value]) => [key, String(value ?? "")]))
      : {};
  } catch {
    return {};
  }
}

async function moveWorkOrderToWeek(db: Awaited<ReturnType<typeof requireDb>>, workOrderId: number, fromWeekId: number, toWeekId: number) {
  if (fromWeekId === toWeekId) return;
  await db.update(otStates).set({ weekId: toWeekId }).where(and(eq(otStates.workOrderId, workOrderId), eq(otStates.weekId, fromWeekId)));
  await db.update(otExecutions).set({ weekId: toWeekId }).where(and(eq(otExecutions.workOrderId, workOrderId), eq(otExecutions.weekId, fromWeekId)));
}

export function buildGuardAssignmentRows(
  weekId: number,
  guard: "A" | "B" | "C" | "D",
  equipmentIds: number[],
  selectedById: number | null
) {
  const validEquipmentIds = equipmentIds.filter(
    (equipmentId): equipmentId is number =>
      Number.isSafeInteger(equipmentId) && equipmentId > 0
  );
  return Array.from(new Set(validEquipmentIds)).map(equipmentId => ({
    weekId,
    guard,
    equipmentId,
    selectedById,
  }));
}

function uniqueCodes(codes: string[]) {
  return Array.from(new Set(codes.map(code => code.trim().toUpperCase()).filter(Boolean)));
}

export const appRouter = router({
  system: systemRouter,
  auth: router({
    me: publicProcedure.query(
      ({ ctx }) =>
        ctx.user ?? (hasSharedAccess(ctx.req) ? createSharedAccessUser(getSharedAccessRole(ctx.req), new Date(), getSharedAccessGuard(ctx.req), getSharedAccessName(ctx.req)) : null)
    ),
    accessUsers: publicProcedure.query(async () => {
      const db = await getDb();
      const managedUsers = db
        ? await db.select({ id: accessCodes.id, label: accessCodes.label, role: accessCodes.role, guard: accessCodes.guard })
          .from(accessCodes)
          .where(eq(accessCodes.active, true))
        : [];
      return [
        { id: "admin", label: "Administrador", role: "admin" as const, guard: "A" as const },
        ...managedUsers.map(user => ({ id: String(user.id), label: user.label, role: user.role, guard: user.guard })),
      ];
    }),
    pinLogin: publicProcedure
      .input(z.object({ pin: z.string().regex(/^\d{4}$/, "El PIN debe tener exactamente 4 dígitos."), selectedUserId: z.string().min(1).max(32).optional() }))
      .mutation(async ({ ctx, input }) => {
        const configuredPin = getConfiguredSharedPin();
        const signingSecret = process.env.JWT_SECRET;
        if (!configuredPin || !signingSecret) {
          throw new TRPCError({
            code: "INTERNAL_SERVER_ERROR",
            message:
              "El acceso por PIN aún no está configurado. Contacta al administrador del sitio.",
          });
        }

        const ip = ctx.req.ip || ctx.req.socket.remoteAddress || "unknown";
        if (!canAttemptSharedPin(ip)) {
          throw new TRPCError({
            code: "TOO_MANY_REQUESTS",
            message:
              "Demasiados intentos. Espera 15 minutos antes de volver a probar.",
          });
        }
        let accessRole: "admin" | "facilitator" | "supervisor" | null = null;
        let accessGuard: "A" | "B" | "C" | "D" = "A";
        let accessName = "Acceso compartido";
        const db = await getDb();
        if (!input.selectedUserId || input.selectedUserId === "admin") {
          if (matchesSharedPin(input.pin, configuredPin)) accessRole = "admin";
        }
        if (!accessRole && db && input.selectedUserId && input.selectedUserId !== "admin") {
          const selectedId = Number(input.selectedUserId);
          const [selected] = Number.isSafeInteger(selectedId)
            ? await db.select().from(accessCodes).where(and(eq(accessCodes.id, selectedId), eq(accessCodes.active, true))).limit(1)
            : [];
          if (selected && matchesAccessCode(input.pin, selected.salt, selected.codeHash)) {
            accessRole = selected.role;
            accessGuard = selected.guard;
            accessName = selected.label;
            await db.update(accessCodes).set({ lastUsedAt: new Date() }).where(eq(accessCodes.id, selected.id));
          }
        } else if (!accessRole && db && !input.selectedUserId) {
          const codes = await db.select().from(accessCodes).where(eq(accessCodes.active, true));
          const matched = codes.find(code => matchesAccessCode(input.pin, code.salt, code.codeHash));
          if (matched) {
            accessRole = matched.role;
            accessGuard = matched.guard;
            accessName = matched.label;
            await db.update(accessCodes).set({ lastUsedAt: new Date() }).where(eq(accessCodes.id, matched.id));
          }
        }
        if (!accessRole) {
          recordFailedSharedPinAttempt(ip);
          throw new TRPCError({ code: "UNAUTHORIZED", message: "Código de acceso incorrecto." });
        }

        clearFailedSharedPinAttempts(ip);
        ctx.res.cookie(
          SHARED_ACCESS_COOKIE,
          createSharedAccessToken(
            signingSecret,
            Date.now() + SHARED_ACCESS_TTL_MS,
            accessRole,
            accessGuard,
            accessName
          ),
          {
            ...getSessionCookieOptions(ctx.req),
            sameSite: "lax",
            maxAge: SHARED_ACCESS_TTL_MS,
          }
        );
        return createSharedAccessUser(accessRole, new Date(), accessGuard, accessName);
      }),
    logout: publicProcedure.mutation(({ ctx }) => {
      const cookieOptions = getSessionCookieOptions(ctx.req);
      ctx.res.clearCookie(COOKIE_NAME, { ...cookieOptions, maxAge: -1 });
      ctx.res.clearCookie(SHARED_ACCESS_COOKIE, {
        ...cookieOptions,
        sameSite: "lax",
        maxAge: -1,
      });
      return { success: true } as const;
    }),
  }),
  maintenance: router({
    bootstrap: sharedAccessProcedure.query(async () => ({
      weeks: await getWeeks(),
      equipment: await getEquipment(),
      warehouses: await getWarehouses(),
      relations: await getGuardRelations(),
    })),

    dashboard: sharedAccessProcedure
      .input(
        z.object({ weekId: z.number().int().positive(), guard: guardSchema })
      )
      .query(({ input, ctx }) => getDashboardData(input.weekId, input.guard, "facilitator", ctx.user?.name ?? "")),

    importWeek: facilitatorProcedure
      .input(
        z.object({
          label: z.string().min(1).max(32),
          weekNumber: z.number().int().min(1).max(53),
          year: z.number().int().min(2020).max(2200),
          guard: guardSchema,
          sourceFileName: z.string().max(255).optional(),
          rows: z
            .array(
              z.object({
                otNumber: z.string().min(1).max(80),
                equipmentCode: z.string().min(1).max(80),
                equipmentDescription: z.string().max(255).optional(),
                description: z.string().max(10000).optional(),
                extra: z.record(z.string(), z.string()).optional(),
              })
            )
            .min(1),
        })
      )
      .mutation(async ({ input, ctx }) => {
        const db = await requireDb();
        const [already] = await db
          .select()
          .from(maintenanceWeeks)
          .where(
            and(
              eq(maintenanceWeeks.year, input.year),
              eq(maintenanceWeeks.weekNumber, input.weekNumber)
            )
          )
          .limit(1);
        const weekId = already?.id ?? insertId(await db.insert(maintenanceWeeks).values({
          label: input.label,
          weekNumber: input.weekNumber,
          year: input.year,
          sourceFileName: input.sourceFileName,
          importedById: ctx.user?.id ?? null,
        }));
        const equipmentCache = new Map<string, number>();
        const assignmentEquipmentIds: number[] = [];
        let updatedCount = 0;
        let duplicateCount = 0;
        let equipmentCount = 0;
        let otCount = 0;
        let purgedCount = 0;

        const incomingOtNumbers = new Set(input.rows.map(row => row.otNumber.trim()));
        if (already) {
          const existingRows = await db
            .select({ workOrder: workOrders, equipmentActive: equipment.activo_en_taller })
            .from(workOrders)
            .innerJoin(equipment, eq(workOrders.equipmentId, equipment.id))
            .where(and(eq(workOrders.weekId, weekId), eq(workOrders.isCancelled, false)));

          for (const existing of existingRows) {
            if (incomingOtNumbers.has(existing.workOrder.otNumber)) continue;
            if (existing.equipmentActive || correctiveSource(existing.workOrder.sourceData)) continue;

            const states = await db
              .select({ status: otStates.status })
              .from(otStates)
              .where(and(eq(otStates.workOrderId, existing.workOrder.id), eq(otStates.weekId, weekId)));
            const canPurge = canPurgeSapRescheduledWorkOrder({
              isInNewExcel: false,
              isCorrective: Boolean(correctiveSource(existing.workOrder.sourceData)),
              equipmentActive: existing.equipmentActive,
              statuses: states.map(state => state.status),
            });
            if (!canPurge) continue;

            const stateIds = await db
              .select({ id: otStates.id })
              .from(otStates)
              .where(and(eq(otStates.workOrderId, existing.workOrder.id), eq(otStates.weekId, weekId)));
            for (const state of stateIds) {
              await db.delete(otStatePendingWarehouses).where(eq(otStatePendingWarehouses.otStateId, state.id));
              await db.delete(otHistory).where(eq(otHistory.otStateId, state.id));
            }
            await db.delete(otStates).where(and(eq(otStates.workOrderId, existing.workOrder.id), eq(otStates.weekId, weekId)));
            await db.delete(otExecutions).where(and(eq(otExecutions.workOrderId, existing.workOrder.id), eq(otExecutions.weekId, weekId)));
            await db.delete(workOrders).where(eq(workOrders.id, existing.workOrder.id));
            purgedCount += 1;
          }
        }

        for (const row of input.rows) {
          const otNumber = row.otNumber.trim();
          const equipmentCode = row.equipmentCode.trim().toUpperCase();
          let equipmentId = equipmentCache.get(equipmentCode);
          if (!equipmentId) {
            const [existingEquipment] = await db.select().from(equipment).where(eq(equipment.code, equipmentCode)).limit(1);
            if (existingEquipment) {
              equipmentId = existingEquipment.id;
              if (row.equipmentDescription?.trim()) {
                await db.update(equipment).set({ description: row.equipmentDescription.trim() }).where(eq(equipment.id, equipmentId));
              }
            } else {
              equipmentId = insertId(await db.insert(equipment).values({ code: equipmentCode, description: row.equipmentDescription || null, activo_en_taller: false }));
              equipmentCount += 1;
            }
            equipmentCache.set(equipmentCode, equipmentId);
          }

          const [existingWorkOrder] = await db.select().from(workOrders).where(eq(workOrders.otNumber, otNumber)).limit(1);
          if (existingWorkOrder) {
            const previousData = parseSourceData(existingWorkOrder.sourceData);
            const technicalData = { ...previousData, ...(row.extra ?? {}), Responsable: ctx.user?.name ?? "Operador logístico" };
            await moveWorkOrderToWeek(db, existingWorkOrder.id, existingWorkOrder.weekId, weekId);
            await db.update(workOrders).set({
              weekId,
              equipmentId,
              description: row.description || existingWorkOrder.description,
              sourceData: JSON.stringify(technicalData),
            }).where(eq(workOrders.id, existingWorkOrder.id));
            updatedCount += 1;
            duplicateCount += 1;
            continue;
          }

          await db.insert(workOrders).values({
            weekId,
            otNumber,
            equipmentId,
            description: row.description || null,
            sourceData: JSON.stringify({ ...(row.extra ?? {}), Responsable: ctx.user?.name ?? "Operador logístico" }),
          });
          assignmentEquipmentIds.push(equipmentId);
          otCount += 1;
        }

        for (const assignment of buildGuardAssignmentRows(weekId, input.guard, assignmentEquipmentIds, ctx.user?.id ?? null)) {
          const [existingAssignment] = await db
            .select({ id: guardAssignments.id })
            .from(guardAssignments)
            .where(and(
              eq(guardAssignments.weekId, assignment.weekId),
              eq(guardAssignments.guard, assignment.guard),
              eq(guardAssignments.equipmentId, assignment.equipmentId),
            ))
            .limit(1);
          if (!existingAssignment) await db.insert(guardAssignments).values(assignment);
        }

        return { weekId, otCount, updatedCount, equipmentCount, duplicateCount, purgedCount };
      }),

    createCorrectiveOt: facilitatorProcedure
      .input(z.object({
        weekId: z.number().int().positive(),
        guard: guardSchema,
        otNumber: z.string().trim().min(1).max(80),
        equipmentCode: z.string().trim().min(1).max(80),
        reservation: z.string().trim().max(160).optional(),
        description: z.string().trim().min(1).max(10000),
      }))
      .mutation(async ({ input, ctx }) => {
        const db = await requireDb();
        const equipmentCode = input.equipmentCode.trim().toUpperCase();
        const [existingEquipment] = await db.select().from(equipment).where(eq(equipment.code, equipmentCode)).limit(1);
        let equipmentId: number;
        if (existingEquipment) {
          equipmentId = existingEquipment.id;
          await db.update(equipment).set({ activo_en_taller: true }).where(eq(equipment.id, equipmentId));
        } else {
          equipmentId = insertId(await db.insert(equipment).values({ code: equipmentCode, activo_en_taller: true }));
        }
        const sourceData = JSON.stringify({
          Reserva: input.reservation ?? "",
          "Descripción de operación": input.description,
          Tipo: "Correctiva",
          Responsable: ctx.user?.name ?? "Operador logístico",
        });
        const [existingWorkOrder] = await db.select().from(workOrders).where(eq(workOrders.otNumber, input.otNumber)).limit(1);
        let workOrderId: number;
        let updated = false;
        if (existingWorkOrder) {
          await moveWorkOrderToWeek(db, existingWorkOrder.id, existingWorkOrder.weekId, input.weekId);
          await db.update(workOrders).set({
            weekId: input.weekId,
            equipmentId,
            description: input.description,
            sourceData,
            isCancelled: false,
            cancelledAt: null,
            cancelledById: null,
            cancellationReason: null,
          }).where(eq(workOrders.id, existingWorkOrder.id));
          workOrderId = existingWorkOrder.id;
          updated = true;
        } else {
          workOrderId = insertId(await db.insert(workOrders).values({
            weekId: input.weekId,
            otNumber: input.otNumber,
            equipmentId,
            description: input.description,
            sourceData,
          }));
        }
        const [assignment] = await db.select({ id: guardAssignments.id })
          .from(guardAssignments)
          .where(and(
            eq(guardAssignments.weekId, input.weekId),
            eq(guardAssignments.guard, input.guard),
            eq(guardAssignments.equipmentId, equipmentId)
          ))
          .limit(1);
        if (!assignment) {
          await db.insert(guardAssignments).values({
            weekId: input.weekId,
            guard: input.guard,
            equipmentId,
            selectedById: ctx.user?.id ?? null,
          });
        }
        return { success: true, updated, workOrderId, equipmentId } as const;
      }),

    updateCorrectiveOt: facilitatorProcedure
      .input(z.object({
        workOrderId: z.number().int().positive(),
        otNumber: z.string().trim().min(1).max(80),
        equipmentCode: z.string().trim().min(1).max(80),
        reservation: z.string().trim().max(160).optional(),
        description: z.string().trim().min(1).max(10000),
      }))
      .mutation(async ({ input, ctx }) => {
        const db = await requireDb();
        const [current] = await db.select().from(workOrders).where(eq(workOrders.id, input.workOrderId)).limit(1);
        if (!current || current.isCancelled || !correctiveSource(current.sourceData)) throw new TRPCError({ code: "NOT_FOUND", message: "Solo se pueden editar OT correctivas manuales activas." });
        const [duplicate] = await db.select().from(workOrders).where(eq(workOrders.otNumber, input.otNumber)).limit(1);
        let targetId = current.id;
        if (duplicate && duplicate.id !== current.id) {
          const sourceStates = await db.select().from(otStates).where(eq(otStates.workOrderId, current.id));
          const targetStates = await db.select().from(otStates).where(eq(otStates.workOrderId, duplicate.id));
          for (const sourceState of sourceStates) {
            const matchingTarget = targetStates.find(state => state.weekId === sourceState.weekId && state.guard === sourceState.guard);
            if (matchingTarget) {
              await db.delete(otStatePendingWarehouses).where(eq(otStatePendingWarehouses.otStateId, sourceState.id));
              await db.delete(otHistory).where(eq(otHistory.otStateId, sourceState.id));
              await db.delete(otStates).where(eq(otStates.id, sourceState.id));
            } else {
              await db.update(otStates).set({ workOrderId: duplicate.id }).where(eq(otStates.id, sourceState.id));
            }
          }
          const sourceExecutions = await db.select().from(otExecutions).where(eq(otExecutions.workOrderId, current.id));
          const targetExecutions = await db.select().from(otExecutions).where(eq(otExecutions.workOrderId, duplicate.id));
          for (const sourceExecution of sourceExecutions) {
            const matchingTarget = targetExecutions.find(execution => execution.weekId === sourceExecution.weekId && execution.guard === sourceExecution.guard);
            if (matchingTarget) await db.delete(otExecutions).where(eq(otExecutions.id, sourceExecution.id));
            else await db.update(otExecutions).set({ workOrderId: duplicate.id }).where(eq(otExecutions.id, sourceExecution.id));
          }
          await db.update(otHistory).set({ workOrderId: duplicate.id }).where(eq(otHistory.workOrderId, current.id));
          await db.delete(workOrders).where(eq(workOrders.id, current.id));
          targetId = duplicate.id;
        }
        const code = input.equipmentCode.trim().toUpperCase();
        const [existingEquipment] = await db.select().from(equipment).where(eq(equipment.code, code)).limit(1);
        const equipmentId = existingEquipment?.id ?? insertId(await db.insert(equipment).values({ code, activo_en_taller: true }));
        if (existingEquipment) await db.update(equipment).set({ activo_en_taller: true }).where(eq(equipment.id, equipmentId));
        await db.update(workOrders).set({ otNumber: input.otNumber, equipmentId, description: input.description, sourceData: JSON.stringify({ Reserva: input.reservation ?? "", "Descripción de operación": input.description, Tipo: "Correctiva", Responsable: ctx.user?.name ?? "Operador logístico" }) }).where(eq(workOrders.id, targetId));
        return { success: true, workOrderId: targetId } as const;
      }),

    createCorrectiveOts: facilitatorProcedure
      .input(z.object({
        weekId: z.number().int().positive(),
        guard: guardSchema,
        rows: z.array(z.object({
          otNumber: z.string().trim().min(1).max(80),
          equipmentCode: z.string().trim().min(1).max(80),
          reservation: z.string().trim().max(160).optional(),
          description: z.string().trim().min(1).max(10000),
        })).min(1).max(200),
      }))
      .mutation(async ({ input, ctx }) => {
        const db = await requireDb();
        const equipmentCache = new Map<string, number>();
        let created = 0;
        let updated = 0;
        let duplicates = 0;
        for (const row of input.rows) {
          const code = row.equipmentCode.trim().toUpperCase();
          let equipmentId = equipmentCache.get(code);
          if (!equipmentId) {
            const [existingEquipment] = await db.select().from(equipment).where(eq(equipment.code, code)).limit(1);
            if (existingEquipment) {
              equipmentId = existingEquipment.id;
              await db.update(equipment).set({ activo_en_taller: true }).where(eq(equipment.id, equipmentId));
            } else {
              equipmentId = insertId(await db.insert(equipment).values({ code, activo_en_taller: true }));
            }
            equipmentCache.set(code, equipmentId);
          }
          const sourceData = JSON.stringify({ Reserva: row.reservation ?? "", "Descripción de operación": row.description, Tipo: "Correctiva", Responsable: ctx.user?.name ?? "Operador logístico" });
          const [existingWorkOrder] = await db.select().from(workOrders).where(eq(workOrders.otNumber, row.otNumber)).limit(1);
          if (existingWorkOrder) {
            await moveWorkOrderToWeek(db, existingWorkOrder.id, existingWorkOrder.weekId, input.weekId);
            await db.update(workOrders).set({ weekId: input.weekId, equipmentId, description: row.description, sourceData, isCancelled: false, cancelledAt: null, cancelledById: null, cancellationReason: null }).where(eq(workOrders.id, existingWorkOrder.id));
            updated += 1;
            duplicates += 1;
          } else {
            await db.insert(workOrders).values({ weekId: input.weekId, otNumber: row.otNumber, equipmentId, description: row.description, sourceData });
            created += 1;
          }
          const [assignment] = await db.select({ id: guardAssignments.id }).from(guardAssignments).where(and(eq(guardAssignments.weekId, input.weekId), eq(guardAssignments.guard, input.guard), eq(guardAssignments.equipmentId, equipmentId))).limit(1);
          if (!assignment) await db.insert(guardAssignments).values({ weekId: input.weekId, guard: input.guard, equipmentId, selectedById: ctx.user?.id ?? null });
        }
        return { success: true, created, updated, duplicates } as const;
      }),

    cancelCorrectiveOt: facilitatorProcedure
      .input(z.object({ workOrderId: z.number().int().positive(), reason: z.string().trim().max(500).optional() }))
      .mutation(async ({ input, ctx }) => {
        const db = await requireDb();
        const [current] = await db.select().from(workOrders).where(eq(workOrders.id, input.workOrderId)).limit(1);
        if (!current || current.isCancelled || !correctiveSource(current.sourceData)) throw new TRPCError({ code: "NOT_FOUND", message: "Solo se pueden anular OT correctivas manuales activas." });
        await db.update(workOrders).set({ isCancelled: true, cancelledAt: new Date(), cancelledById: ctx.user?.id ?? null, cancellationReason: input.reason || "Anulada por el Facilitador" }).where(eq(workOrders.id, input.workOrderId));
        return { success: true };
      }),

    saveAssignments: facilitatorProcedure
      .input(
        z.object({
          weekId: z.number().int().positive(),
          guard: guardSchema,
          equipmentIds: z.array(z.number().int().positive()),
        })
      )
      .mutation(async ({ input, ctx }) => {
        const db = await requireDb();
        const weekEquipment = await db.select({ id: equipment.id })
          .from(equipment)
          .innerJoin(workOrders, eq(workOrders.equipmentId, equipment.id))
          .where(and(eq(workOrders.weekId, input.weekId), eq(workOrders.isCancelled, false)));
        const allowedEquipmentIds = new Set(weekEquipment.map(item => item.id));
        if (input.equipmentIds.some(equipmentId => !allowedEquipmentIds.has(equipmentId))) {
          throw new TRPCError({ code: "BAD_REQUEST", message: "Solo puedes asignar equipos de la semana seleccionada." });
        }
        await db
          .delete(guardAssignments)
          .where(
            and(
              eq(guardAssignments.weekId, input.weekId),
              eq(guardAssignments.guard, input.guard)
            )
          );
        if (input.equipmentIds.length) {
          await db.insert(guardAssignments).values(
            input.equipmentIds.map(equipmentId => ({
              weekId: input.weekId,
              guard: input.guard,
              equipmentId,
              selectedById: ctx.user?.id ?? null,
            }))
          );
        }
        return getDashboardData(input.weekId, input.guard);
      }),

    setEquipmentWorkshopStatus: facilitatorProcedure
      .input(z.object({ equipmentId: z.number().int().positive(), active: z.boolean() }))
      .mutation(async ({ input, ctx }) => {
        const db = await requireDb();
        const [item] = await db.select({ id: equipment.id }).from(equipment).where(eq(equipment.id, input.equipmentId)).limit(1);
        if (!item) throw new TRPCError({ code: "NOT_FOUND", message: "Equipo no encontrado." });
        await db.update(equipment).set({ activo_en_taller: input.active }).where(eq(equipment.id, input.equipmentId));
        const relatedOrders = await db.select({ id: workOrders.id, sourceData: workOrders.sourceData }).from(workOrders).where(eq(workOrders.equipmentId, input.equipmentId));
        for (const order of relatedOrders) {
          let data: Record<string, unknown> = {};
          try { data = order.sourceData ? JSON.parse(order.sourceData) : {}; } catch { data = {}; }
          data["Última acción de equipo"] = input.active ? "Ingresar a Guardia" : "Liberar Equipo";
          data.Responsable = ctx.user?.name ?? "Operador logístico";
          await db.update(workOrders).set({ sourceData: JSON.stringify(data) }).where(eq(workOrders.id, order.id));
        }
        if (!input.active) {
          await db.delete(guardAssignments).where(eq(guardAssignments.equipmentId, input.equipmentId));
          await db.delete(supervisorAssignments).where(eq(supervisorAssignments.equipmentId, input.equipmentId));
        }
        return { success: true, active: input.active } as const;
      }),

    saveSupervisorAssignments: supervisorProcedure
      .input(z.object({ weekId: z.number().int().positive(), guard: guardSchema, equipmentIds: z.array(z.number().int().positive()) }))
      .mutation(async ({ input, ctx }) => {
        const db = await requireDb();
        const weekEquipment = await db.select({ id: equipment.id })
          .from(equipment)
          .innerJoin(workOrders, eq(workOrders.equipmentId, equipment.id))
          .where(and(eq(workOrders.weekId, input.weekId), eq(workOrders.isCancelled, false)));
        const allowedEquipmentIds = new Set(weekEquipment.map(item => item.id));
        if (input.equipmentIds.some(equipmentId => !allowedEquipmentIds.has(equipmentId))) {
          throw new TRPCError({ code: "BAD_REQUEST", message: "Solo puedes seleccionar equipos de la semana activa." });
        }
        const supervisorName = "Facilitador logístico";
        await db.delete(supervisorAssignments).where(and(
          eq(supervisorAssignments.supervisorName, supervisorName),
          eq(supervisorAssignments.weekId, input.weekId),
          eq(supervisorAssignments.guard, input.guard),
        ));
        if (input.equipmentIds.length) {
          await db.insert(supervisorAssignments).values(input.equipmentIds.map(equipmentId => ({
            supervisorName, weekId: input.weekId, guard: input.guard, equipmentId,
          })));
        }
        return getDashboardData(input.weekId, input.guard, "supervisor", supervisorName);
      }),

    updateReception: facilitatorProcedure
      .input(
        z.object({
          workOrderId: z.number().int().positive(),
          weekId: z.number().int().positive(),
          guard: guardSchema,
          status: statusSchema,
          warehouseCodes: z.array(z.string().min(1).max(64)).default([]),
          responsibleName: z.string().max(160).optional(),
          note: z.string().max(10000).optional(),
        })
      )
      .mutation(async ({ input, ctx }) => {
        const db = await requireDb();
        const [workOrder] = await db
          .select()
          .from(workOrders)
          .where(eq(workOrders.id, input.workOrderId))
          .limit(1);
        if (!workOrder || workOrder.weekId !== input.weekId)
          throw new TRPCError({ code: "NOT_FOUND", message: "OT no encontrada para esta semana." });
        if (!hasReservation(workOrder.sourceData))
          throw new TRPCError({
            code: "BAD_REQUEST",
            message: "OT sin reserva: no requiere clasificación logística de repuestos.",
          });
        const warehouseCodes = uniqueCodes(input.warehouseCodes);
        const validationMessage = validateReceptionInput(input.status, warehouseCodes);
        if (validationMessage) throw new TRPCError({ code: "BAD_REQUEST", message: validationMessage });
        const selectedWarehouses = warehouseCodes.length
          ? await db.select().from(warehouses).where(and(inArray(warehouses.code, warehouseCodes), eq(warehouses.active, true)))
          : [];
        if (selectedWarehouses.length !== warehouseCodes.length)
          throw new TRPCError({ code: "BAD_REQUEST", message: "Uno o más almacenes no están en el catálogo activo." });
        const [existing] = await db.select().from(otStates).where(and(eq(otStates.workOrderId, input.workOrderId), eq(otStates.weekId, input.weekId))).limit(1);
        if (existing && !canReviewReception(existing.status)) {
          throw new TRPCError({
            code: "CONFLICT",
            message: "Esta OT ya fue recibida/revisada y no se puede revisar nuevamente.",
          });
        }
        const responsibleName = input.responsibleName?.trim() || ctx.user?.name || ctx.user?.email || "Usuario actual";
        const newPendingWarehouse = warehouseCodes.join(", ") || null;
        let stateId: number;
        await db.transaction(async tx => {
          const inserted = await tx.insert(otStates).values({ workOrderId: input.workOrderId, weekId: input.weekId, guard: input.guard, status: input.status, pendingWarehouse1: false, pendingWarehouse2: false, responsibleName, reviewedByRole: "facilitator", note: input.note?.trim() || null, updatedById: ctx.user?.id ?? null });
          stateId = insertId(inserted);
          if (selectedWarehouses.length) await tx.insert(otStatePendingWarehouses).values(selectedWarehouses.map(warehouse => ({ otStateId: stateId, warehouseId: warehouse.id, warehouseCodeSnapshot: warehouse.code, warehouseNameSnapshot: warehouse.name })));
          await tx.insert(otHistory).values({ otStateId: stateId, workOrderId: input.workOrderId, weekId: input.weekId, guard: input.guard, previousStatus: null, newStatus: input.status, previousPendingWarehouse: null, newPendingWarehouse, note: input.note?.trim() || null, actorUserId: ctx.user?.id ?? null, actorName: ctx.user?.name || ctx.user?.email || "Usuario actual" });
        });
        return { success: true, stateId: stateId! };
      }),

    markOperationalExecution: operationalProcedure
      .input(
        z.object({
          workOrderId: z.number().int().positive(),
          weekId: z.number().int().positive(),
          guard: guardSchema,
          closed: z.boolean(),
        })
      )
      .mutation(async ({ input, ctx }) => {
        const db = await requireDb();
        const [workOrder] = await db
          .select()
          .from(workOrders)
          .where(eq(workOrders.id, input.workOrderId))
          .limit(1);
        if (!workOrder || workOrder.weekId !== input.weekId)
          throw new TRPCError({ code: "NOT_FOUND", message: "OT no encontrada para esta semana." });
        const status = input.closed ? "CERRADA_NOTIFICADA" : "ABIERTA";
        const values = {
          status,
          executedById: ctx.user?.id ?? null,
          executedByName: ctx.user?.name || ctx.user?.email || "Usuario actual",
          executedAt: new Date(),
        } as const;
        const [existing] = await db
          .select()
          .from(otExecutions)
          .where(and(eq(otExecutions.workOrderId, input.workOrderId), eq(otExecutions.weekId, input.weekId), eq(otExecutions.guard, input.guard)))
          .limit(1);
        if (existing) await db.update(otExecutions).set(values).where(eq(otExecutions.id, existing.id));
        else await db.insert(otExecutions).values({
          workOrderId: input.workOrderId,
          weekId: input.weekId,
          guard: input.guard,
          ...values,
        });
        let equipmentDeactivated = false;
        if (input.closed) {
          const relatedOrders = await db.select({ id: workOrders.id })
            .from(workOrders)
            .where(and(eq(workOrders.weekId, input.weekId), eq(workOrders.equipmentId, workOrder.equipmentId), eq(workOrders.isCancelled, false)));
          const relatedIds = relatedOrders.map(item => item.id);
          const relatedExecutions = relatedIds.length
            ? await db.select({ status: otExecutions.status })
              .from(otExecutions)
              .where(and(eq(otExecutions.weekId, input.weekId), eq(otExecutions.guard, input.guard), inArray(otExecutions.workOrderId, relatedIds)))
            : [];
          if (relatedOrders.length && relatedExecutions.length === relatedOrders.length && areAllWorkOrdersClosed(relatedExecutions.map(item => item.status))) {
            await db.update(equipment).set({ activo_en_taller: false }).where(eq(equipment.id, workOrder.equipmentId));
            await db.delete(guardAssignments).where(eq(guardAssignments.equipmentId, workOrder.equipmentId));
            await db.delete(supervisorAssignments).where(eq(supervisorAssignments.equipmentId, workOrder.equipmentId));
            equipmentDeactivated = true;
          }
        }
        return { success: true, status, equipmentDeactivated };
      }),

    history: sharedAccessProcedure
      .input(z.object({ workOrderId: z.number().int().positive() }))
      .query(({ input }) => getOtHistory(input.workOrderId)),
    historicalByEquipment: sharedAccessProcedure
      .input(z.object({ equipmentQuery: z.string().trim().min(1).max(80), otType: z.enum(["TODAS", "PROGRAMADAS", "CORRECTIVAS"]).default("TODAS") }))
      .query(({ input }) => searchHistoricalWorkOrders(input.equipmentQuery, input.otType)),

    generateHandover: facilitatorProcedure
      .input(z.object({ weekId: z.number().int().positive(), fromGuard: guardSchema, toGuard: guardSchema }))
      .mutation(async ({ input, ctx }) => {
        if (input.fromGuard === input.toGuard) throw new TRPCError({ code: "BAD_REQUEST", message: "La guardia que entrega y recibe deben ser diferentes." });
        const db = await requireDb();
        const dashboard = await getDashboardData(input.weekId, input.fromGuard);
        const pendingRows = dashboard.rows.filter(
          row =>
            hasReservation(row.workOrder.sourceData) &&
            row.state &&
            (row.state.status === "PARCIAL" || row.state.status === "FALTA_RECIBIR")
        );
        return db.transaction(async tx => {
          const [already] = await tx.select().from(handovers).where(and(eq(handovers.weekId, input.weekId), eq(handovers.fromGuard, input.fromGuard), eq(handovers.toGuard, input.toGuard), inArray(handovers.status, ["ENVIADO", "CONFIRMADO"]))).limit(1);
          if (already) throw new TRPCError({ code: "CONFLICT", message: "Este relevo ya fue generado o confirmado." });
          const inserted = await tx.insert(handovers).values({ weekId: input.weekId, fromGuard: input.fromGuard, toGuard: input.toGuard, status: "ENVIADO", createdById: ctx.user?.id ?? null });
          const handoverId = insertId(inserted);
          for (const row of pendingRows) {
            const state = row.state!;
            const pendingWarehouses = row.pendingWarehouses.map(item => ({ code: item.warehouseCodeSnapshot, name: item.warehouseNameSnapshot }));
            await tx.insert(handoverItems).values({ handoverId, otStateId: state.id, workOrderId: row.workOrder.id, statusSnapshot: state.status, pendingWarehousesSnapshot: JSON.stringify(pendingWarehouses), fromGuardSnapshot: input.fromGuard, toGuardSnapshot: input.toGuard, createdByIdSnapshot: ctx.user?.id ?? null, createdAtSnapshot: new Date() });
            const [incoming] = await tx.select().from(otStates).where(and(eq(otStates.workOrderId, row.workOrder.id), eq(otStates.weekId, input.weekId), eq(otStates.guard, input.toGuard))).limit(1);
            let incomingId = incoming?.id;
            if (incoming) await tx.update(otStates).set({ status: state.status, pendingWarehouse1: false, pendingWarehouse2: false, inheritedFromHandoverId: handoverId, responsibleName: state.responsibleName, note: state.note, updatedById: ctx.user?.id ?? null, updatedAt: new Date() }).where(eq(otStates.id, incoming.id));
            else { const created = await tx.insert(otStates).values({ workOrderId: row.workOrder.id, weekId: input.weekId, guard: input.toGuard, status: state.status, pendingWarehouse1: false, pendingWarehouse2: false, inheritedFromHandoverId: handoverId, responsibleName: state.responsibleName, note: state.note, updatedById: ctx.user?.id ?? null }); incomingId = insertId(created); }
            await tx.delete(otStatePendingWarehouses).where(eq(otStatePendingWarehouses.otStateId, incomingId!));
            if (row.pendingWarehouses.length) await tx.insert(otStatePendingWarehouses).values(row.pendingWarehouses.map(item => ({ otStateId: incomingId!, warehouseId: item.warehouseId, warehouseCodeSnapshot: item.warehouseCodeSnapshot, warehouseNameSnapshot: item.warehouseNameSnapshot })));
            const [assignment] = await tx.select().from(guardAssignments).where(and(eq(guardAssignments.weekId, input.weekId), eq(guardAssignments.guard, input.toGuard), eq(guardAssignments.equipmentId, row.workOrder.equipmentId))).limit(1);
            if (!assignment) await tx.insert(guardAssignments).values({ weekId: input.weekId, guard: input.toGuard, equipmentId: row.workOrder.equipmentId, selectedById: ctx.user?.id ?? null });
          }
          return { handoverId, pendingCount: pendingRows.length };
        });
      }),

    latestHandover: sharedAccessProcedure
      .input(
        z.object({
          weekId: z.number().int().positive(),
          fromGuard: guardSchema,
          toGuard: guardSchema,
        })
      )
      .query(({ input }) =>
        getLatestHandover(input.weekId, input.fromGuard, input.toGuard)
      ),

    confirmHandover: facilitatorProcedure
      .input(z.object({ handoverId: z.number().int().positive() }))
      .mutation(async ({ input, ctx }) => {
        const db = await requireDb();
        await db
          .update(handovers)
          .set({
            status: "CONFIRMADO",
            confirmedById: ctx.user?.id ?? null,
            confirmedAt: new Date(),
          })
          .where(eq(handovers.id, input.handoverId));
        return { success: true };
      }),

    saveRelations: adminProcedure
      .input(
        z.object({
          relations: z.array(
            z.object({ fromGuard: guardSchema, toGuard: guardSchema })
          ),
        })
      )
      .mutation(async ({ input, ctx }) => {
        const db = await requireDb();
        for (const relation of input.relations) {
          if (relation.fromGuard === relation.toGuard) continue;
          const [existing] = await db
            .select()
            .from(guardRelations)
            .where(eq(guardRelations.fromGuard, relation.fromGuard))
            .limit(1);
          if (existing) {
            await db
              .update(guardRelations)
              .set({
                toGuard: relation.toGuard,
                updatedById: ctx.user?.id ?? null,
                updatedAt: new Date(),
              })
              .where(eq(guardRelations.id, existing.id));
          } else {
            await db.insert(guardRelations).values({
              fromGuard: relation.fromGuard,
              toGuard: relation.toGuard,
              updatedById: ctx.user?.id ?? null,
            });
          }
        }
        return getGuardRelations();
      }),

    listUsers: adminProcedure.query(async () => {
      const db = await requireDb();
      return db
        .select({
          id: users.id,
          name: users.name,
          email: users.email,
          role: users.role,
          lastSignedIn: users.lastSignedIn,
        })
        .from(users)
        .orderBy(desc(users.lastSignedIn));
    }),

    listAccessCodes: adminProcedure.query(async () => {
      const db = await requireDb();
      return db
        .select({ id: accessCodes.id, label: accessCodes.label, role: accessCodes.role, guard: accessCodes.guard, active: accessCodes.active, createdAt: accessCodes.createdAt, lastUsedAt: accessCodes.lastUsedAt })
        .from(accessCodes)
        .orderBy(desc(accessCodes.createdAt));
    }),

    createAccessCode: adminProcedure
      .input(z.object({ label: z.string().trim().min(2).max(120), role: z.literal("facilitator"), guard: guardSchema, pin: z.string().regex(/^\d{4}$/, "El PIN debe tener exactamente 4 dígitos.") }))
      .mutation(async ({ input, ctx }) => {
        const db = await requireDb();
        const [existingGuard] = await db.select({ id: accessCodes.id }).from(accessCodes).where(and(eq(accessCodes.guard, input.guard), eq(accessCodes.active, true))).limit(1);
        if (existingGuard) throw new TRPCError({ code: "CONFLICT", message: `Ya existe un facilitador activo para la Guardia ${input.guard === "A" ? "1" : input.guard === "B" ? "2" : input.guard === "C" ? "3" : "4"}.` });
        const { salt, hash } = hashAccessCode(input.pin);
        const inserted = await db.insert(accessCodes).values({ label: input.label, role: input.role, guard: input.guard, salt, codeHash: hash, createdById: ctx.user?.id ?? null });
        return { success: true, id: insertId(inserted) };
      }),

    updateAccessCode: adminProcedure
      .input(z.object({ id: z.number().int().positive(), label: z.string().trim().min(2).max(120), guard: guardSchema, pin: z.string().regex(/^\d{4}$/, "El PIN debe tener exactamente 4 dígitos.") }))
      .mutation(async ({ input }) => {
        const db = await requireDb();
        const [current] = await db.select({ id: accessCodes.id, role: accessCodes.role, active: accessCodes.active }).from(accessCodes).where(eq(accessCodes.id, input.id)).limit(1);
        if (!current || !current.active || current.role !== "facilitator") throw new TRPCError({ code: "NOT_FOUND", message: "Facilitador no encontrado." });
        const [existingGuard] = await db.select({ id: accessCodes.id }).from(accessCodes).where(and(eq(accessCodes.guard, input.guard), eq(accessCodes.active, true))).limit(1);
        if (existingGuard && existingGuard.id !== input.id) throw new TRPCError({ code: "CONFLICT", message: "Ya existe otro facilitador activo asignado a esa guardia." });
        const { salt, hash } = hashAccessCode(input.pin);
        await db.update(accessCodes).set({ label: input.label, guard: input.guard, salt, codeHash: hash }).where(eq(accessCodes.id, input.id));
        return { success: true };
      }),

    deleteAccessCode: adminProcedure
      .input(z.object({ id: z.number().int().positive() }))
      .mutation(async ({ input }) => {
        const db = await requireDb();
        const [current] = await db.select({ id: accessCodes.id, role: accessCodes.role }).from(accessCodes).where(eq(accessCodes.id, input.id)).limit(1);
        if (!current || current.role !== "facilitator") throw new TRPCError({ code: "NOT_FOUND", message: "Facilitador no encontrado." });
        await db.delete(accessCodes).where(eq(accessCodes.id, input.id));
        return { success: true };
      }),

    revokeAccessCode: adminProcedure
      .input(z.object({ id: z.number().int().positive() }))
      .mutation(async ({ input }) => {
        const db = await requireDb();
        await db.update(accessCodes).set({ active: false }).where(eq(accessCodes.id, input.id));
        return { success: true };
      }),

    updateAccessCodeGuard: adminProcedure
      .input(z.object({ id: z.number().int().positive(), guard: guardSchema }))
      .mutation(async ({ input }) => {
        const db = await requireDb();
        await db.update(accessCodes).set({ guard: input.guard }).where(eq(accessCodes.id, input.id));
        return { success: true };
      }),

    updateUserRole: adminProcedure
      .input(z.object({ userId: z.number().int().positive(), role: managedRoleSchema }))
      .mutation(async ({ input }) => {
        const db = await requireDb();
        const [target] = await db.select({ openId: users.openId }).from(users).where(eq(users.id, input.userId)).limit(1);
        if (!target) throw new TRPCError({ code: "NOT_FOUND", message: "Usuario no encontrado." });
        if (target.openId === ENV.ownerOpenId && input.role !== "admin") {
          throw new TRPCError({ code: "FORBIDDEN", message: "El propietario debe conservar el perfil Administrador." });
        }
        await db.update(users).set({ role: input.role }).where(eq(users.id, input.userId));
        return { success: true };
      }),
  }),
});

export type AppRouter = typeof appRouter;
