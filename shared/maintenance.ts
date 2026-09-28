export type ReceptionStatus = "SIN_REVISAR" | "RECIBIDA" | "PARCIAL" | "FALTA_RECIBIR";

export const WAREHOUSE_CODES = ["CENTRAL", "FERREYROS", "DONALDSON"] as const;
export type WarehouseCode = string;

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
