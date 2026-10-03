export type ExcelImportRow = {
  otNumber: string;
  equipmentCode: string;
  equipmentDescription?: string;
  description?: string;
  extra?: Record<string, string>;
  rowNumber?: number;
};

export type ExcelImportInvalidRow = {
  rowNumber: number;
  otNumber: string;
  reason: string;
};

export function normalizeExcelHeader(value: string) {
  return value
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]/g, "");
}

function findColumn(keys: string[], aliases: string[]) {
  const normalized = keys.map(key => ({
    key,
    value: normalizeExcelHeader(key),
  }));
  for (const alias of aliases) {
    const exact = normalized.find(item => item.value === alias);
    if (exact) return exact.key;
  }
  for (const alias of aliases) {
    if (alias.length < 4) continue;
    const partial = normalized.find(item => item.value.includes(alias));
    if (partial) return partial.key;
  }
  return undefined;
}

function normalizeEquipmentCode(value: unknown, sourceKey: string) {
  const code = String(value ?? "").trim().toUpperCase();
  const source = normalizeExcelHeader(sourceKey);
  if (source === "descripciondeubicaciontecnica" || source === "ubicaciontecnica") {
    if (code === "0") return "";
    return code.replace(/^EQ_O/, "");
  }
  return code;
}

/**
 * Converts an Excel sheet matrix into records using its actual header row.
 * Maintenance plans often have title/metadata rows above the headers and a
 * weekday row immediately above the numbered daily columns.
 */
export function rowsFromExcelMatrix(matrix: unknown[][]) {
  const headerIndex = matrix.findIndex(row => {
    const keys = row.map(value => normalizeExcelHeader(String(value ?? "")));
    const hasOt = keys.some(key =>
      ["ordendetrabajo", "numeroot", "nroot", "ot", "workorder"].includes(key)
    );
    const hasEquipment = keys.some(key =>
      [
        "descripciondeubicaciontecnica",
        "ubicaciontecnica",
        "codigodeequipo",
        "codigoequipo",
        "equipo",
        "tag",
        "asset",
      ].includes(key)
    );
    return hasOt && hasEquipment;
  });
  if (headerIndex < 0) return null;

  const headerRow = matrix[headerIndex] ?? [];
  const previousRow = matrix[headerIndex - 1] ?? [];
  const headers = headerRow.map((value, index) => {
    let label = String(value ?? "")
      .replace(/\s+/g, " ")
      .trim();
    const previous = String(previousRow[index] ?? "")
      .replace(/\s+/g, " ")
      .trim();
    if (!label && previous) label = previous;
    else if (/^\d+$/.test(label))
      label = previous ? `${previous} ${label}` : "";
    return label;
  });

  return matrix.slice(headerIndex + 1).flatMap((values, offset) => {
    if (!values.some(value => String(value ?? "").trim())) return [];
    const record: Record<string, unknown> = {};
    headers.forEach((header, index) => {
      if (header) record[header] = values[index] ?? "";
    });
    Object.defineProperty(record, "__excelRowNumber", {
      value: headerIndex + offset + 2,
      enumerable: false,
    });
    return [record];
  });
}

function mergeDistinctValues(previous: string, next: string) {
  const toValues = (value: string) => {
    try {
      const parsed: unknown = JSON.parse(value);
      if (Array.isArray(parsed)) return parsed.map(item => String(item));
    } catch {
      // A normal spreadsheet cell is kept as a single value.
    }
    return value ? [value] : [];
  };
  const values = [...toValues(previous), ...toValues(next)].filter(
    (value, index, all) => all.indexOf(value) === index
  );
  if (values.length <= 1) return values[0] ?? "";
  return JSON.stringify(values);
}

export function mapExcelRows(rawRows: Record<string, unknown>[]) {
  if (!rawRows.length) throw new Error("No se encontraron filas con datos.");
  const keys = Object.keys(rawRows[0] ?? {});
  const otKey = findColumn(keys, [
    "ordendetrabajo",
    "numeroot",
    "nroot",
    "ot",
    "workorder",
  ]);
  const equipmentKey = findColumn(keys, [
    "descripciondeubicaciontecnica",
    "ubicaciontecnica",
    "codigodeequipo",
    "codigoequipo",
    "equipmentcode",
    "equipo",
    "asset",
    "tag",
  ]);
  const descriptionKey = findColumn(keys, [
    "descripciondeordendetrabajo",
    "descripciondeot",
    "descripcionot",
    "shorttext",
    "descripcion",
    "tarea",
    "trabajo",
    "texto",
    "denominacion",
  ]);
  const equipmentDescriptionKey = findColumn(keys, [
    "descripciondeequipo",
    "equipmentdescription",
    "nombreequipo",
  ]);
  const operationKey = findColumn(keys, [
    "numerooperacion",
    "noperacion",
    "operationnumber",
    "operacion",
  ]);
  if (!otKey || !equipmentKey) {
    throw new Error(
      "No pude detectar las columnas de OT y equipo. Usa encabezados como OT, Orden de trabajo, Equipo, TAG o Descripción de ubicación técnica."
    );
  }

  const mapped = rawRows.map((raw, index) => {
    const extra: Record<string, string> = {};
    for (const key of keys) extra[key] = String(raw[key] ?? "").trim();
    return {
      otNumber: String(raw[otKey] ?? "").trim(),
      equipmentCode: normalizeEquipmentCode(raw[equipmentKey], equipmentKey),
      equipmentDescription: equipmentDescriptionKey
        ? String(raw[equipmentDescriptionKey] ?? "").trim()
        : "",
      description: descriptionKey
        ? String(raw[descriptionKey] ?? "").trim()
        : "",
      extra,
      rowNumber: Number((raw as { __excelRowNumber?: number }).__excelRowNumber) || index + 2,
    } satisfies ExcelImportRow;
  });

  const candidates = mapped.filter(row => row.otNumber || row.equipmentCode);
  const invalidRows = candidates.flatMap(row => {
    if (!row.otNumber && !row.equipmentCode) return [{ rowNumber: row.rowNumber, otNumber: "", reason: "Falta Número de OT y código de Equipo." }];
    if (!row.otNumber) return [{ rowNumber: row.rowNumber, otNumber: "", reason: "Falta Número de OT." }];
    if (!row.equipmentCode) return [{ rowNumber: row.rowNumber, otNumber: row.otNumber, reason: "Falta código de Equipo." }];
    return [];
  });
  const invalid = invalidRows.length;
  const valid = mapped.filter(row => row.otNumber && row.equipmentCode);
  const seen = new Set<string>();
  let duplicates = 0;
  for (const row of valid) {
    if (seen.has(row.otNumber)) duplicates += 1;
    seen.add(row.otNumber);
  }

  // The weekly maintenance workbook contains one row per operation, while the
  // application tracks reception and handovers at the OT level. Preserve each
  // operation's spreadsheet cells as distinct values on the single OT record.
  const rows = operationKey
    ? valid.reduce<ExcelImportRow[]>((grouped, row) => {
        const existing = grouped.find(item => item.otNumber === row.otNumber);
        if (!existing) return [...grouped, row];
        const merged = { ...(existing.extra ?? {}) };
        for (const key of keys)
          merged[key] = mergeDistinctValues(
            merged[key] ?? "",
            row.extra?.[key] ?? ""
          );
        existing.extra = merged;
        existing.description = mergeDistinctValues(
          existing.description ?? "",
          row.description ?? ""
        );
        return grouped;
      }, [])
    : valid;

  return {
    rows: rows.map(({ rowNumber: _rowNumber, ...row }) => row),
    invalid,
    invalidRows,
    duplicates,
    detectedColumns: {
      ot: otKey,
      equipment: equipmentKey,
      description: descriptionKey ?? null,
    },
  };
}
