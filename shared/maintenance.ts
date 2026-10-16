export type ReceptionStatus = "SIN_REVISAR" | "RECIBIDA" | "PARCIAL" | "FALTA_RECIBIR" | "REVISADO";
export type ReceptionAction = "REVIEW" | "COMPLETED" | "RESOLVE_DIFFERENCES";

export const WAREHOUSE_CODES = ["CENTRAL", "FERREYROS", "DONALDSON"] as const;
export type WarehouseCode = string;

function normalizeSourceKey(value: string) {
  return value
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]/g, "");
}

export function getReservationValue(sourceData?: string | null) {
  if (!sourceData) return "";
  try {
    const parsed: unknown = JSON.parse(sourceData);
    if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) return "";
    const entry = Object.entries(parsed as Record<string, unknown>).find(
      ([key]) => normalizeSourceKey(key) === "reserva"
    )?.[1];
    const value = String(entry ?? "").trim();
    return value === "0" || value === "—" || value === "-" ? "" : value;
  } catch {
    return "";
  }
}

export function hasReservation(sourceData?: string | null) {
  return Boolean(getReservationValue(sourceData));
}

export function isCorrectiveWorkOrder(sourceData?: string | null) {
  if (!sourceData) return false;
  try {
    const parsed = JSON.parse(sourceData) as Record<string, unknown>;
    return String(parsed.Tipo ?? "").trim().toLowerCase() === "correctiva";
  } catch {
    return false;
  }
}

export function filterActiveWorkshopEquipment<T extends { activo_en_taller?: boolean | null }>(equipment: T[]) {
  return equipment.filter(item => item.activo_en_taller === true);
}

export function isWorkOrderClosed(status?: string | null) {
  return status === "CERRADA_NOTIFICADA";
}

export function canCloseWorkOrder(input: { executionStatus?: string | null; receptionStatus?: string | null; requiresLogistics?: boolean; isCorrective?: boolean }) {
  if (isWorkOrderClosed(input.executionStatus)) return false;
  if (input.isCorrective || !input.requiresLogistics) return true;
  return input.receptionStatus === "REVISADO";
}

export function areAllWorkOrdersClosed(statuses: Array<string | null | undefined>) {
  return statuses.length > 0 && statuses.every(isWorkOrderClosed);
}

export type EquipmentCycleFilter = "TODOS" | "PENDIENTES" | "FINALIZADOS";
export type OtCycleFilter = "TODAS" | "SIN_REVISAR" | "PENDIENTES" | "REVISADAS" | "FINALIZADAS";

export function isEquipmentCycleFinalized(statuses: Array<string | null | undefined>) {
  return areAllWorkOrdersClosed(statuses);
}

export function matchesOtCycleFilter(input: { receptionStatus?: string | null; executionStatus?: string | null; hasReceptionState?: boolean; isCorrective?: boolean }, filter: OtCycleFilter) {
  const finalized = isWorkOrderClosed(input.executionStatus);
  if (filter === "TODAS") return true;
  if (filter === "FINALIZADAS") return finalized;
  if (input.isCorrective) return false;
  if (finalized) return false;
  if (filter === "SIN_REVISAR") return !input.hasReceptionState || !input.receptionStatus || input.receptionStatus === "SIN_REVISAR";
  if (filter === "PENDIENTES") return input.receptionStatus === "PARCIAL" || input.receptionStatus === "FALTA_RECIBIR";
  return input.receptionStatus === "RECIBIDA" || input.receptionStatus === "REVISADO";
}

export function canPurgeSapRescheduledWorkOrder(input: {
  isInNewExcel: boolean;
  isCorrective: boolean;
  equipmentActive: boolean;
  statuses: Array<string | null | undefined>;
}) {
  if (input.isInNewExcel || input.isCorrective || input.equipmentActive) return false;
  return input.statuses.every(status => status === "SIN_REVISAR");
}

export function canReviewReception(status?: ReceptionStatus | null) {
  return true;
}

export function getReceptionAction(status?: ReceptionStatus | null): ReceptionAction {
  if (status === "REVISADO") return "COMPLETED";
  if (status === "PARCIAL" || status === "FALTA_RECIBIR") return "RESOLVE_DIFFERENCES";
  return "REVIEW";
}

export function getPendingWarehouseLabel(codes: string[] | boolean | null | undefined, legacySecond?: boolean) {
  if (typeof codes === "boolean") {
    if (codes && legacySecond) return "AMBOS";
    if (codes) return "ALMACEN_1";
    if (legacySecond) return "ALMACEN_2";
    return null;
  }
  if (!codes?.length) return null;
  return codes.join(", ");
}

export function validateReceptionInput(status: ReceptionStatus, warehouseCodes: string[] | boolean, legacySecond = false) {
  const codes = typeof warehouseCodes === "boolean"
    ? (warehouseCodes && legacySecond ? ["LEGACY_1", "LEGACY_2"] : warehouseCodes ? ["LEGACY_1"] : legacySecond ? ["LEGACY_2"] : [])
    : Array.from(new Set(warehouseCodes.map(code => code.trim().toUpperCase()).filter(Boolean)));
  if (status === "RECIBIDA" && codes.length) return "Una OT Completa no puede tener almacenes pendientes.";
  if ((status === "PARCIAL" || status === "FALTA_RECIBIR") && !codes.length) return "Selecciona al menos un almacén pendiente.";
  return null;
}

export function calculateReceptionPercent(total: number, received: number) {
  if (total <= 0) return 0;
  return Math.round((Math.max(0, Math.min(received, total)) / total) * 100);
}
