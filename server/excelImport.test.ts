import { describe, expect, it } from "vitest";
import { mapExcelRows, rowsFromExcelMatrix } from "@shared/excelImport";
import { buildGuardAssignmentRows, buildGuardWorkOrderAssignmentRows, insertId } from "./routers";

describe("Excel import mapping", () => {
  it("reads MySQL insert IDs from both direct and tuple results", () => {
    expect(insertId({ insertId: 7 })).toBe(7);
    expect(insertId([{ insertId: "8" }, []])).toBe(8);
  });

  it("assigns each imported equipment once and ignores invalid IDs", () => {
    expect(buildGuardAssignmentRows(1, "A", [0, -1, 3, 3, 7], 22)).toEqual([
      { weekId: 1, guard: "A", equipmentId: 3, selectedById: 22 },
      { weekId: 1, guard: "A", equipmentId: 7, selectedById: 22 },
    ]);
  });

  it("assigns selected work orders once without creating equipment copies", () => {
    expect(buildGuardWorkOrderAssignmentRows(4, "C", [101, 101, 0, 205], 33)).toEqual([
      { weekId: 4, guard: "C", workOrderId: 101, selectedById: 33 },
      { weekId: 4, guard: "C", workOrderId: 205, selectedById: 33 },
    ]);
  });

  it("detects common Spanish headers and keeps original columns", () => {
    const result = mapExcelRows([
      {
        "Orden de trabajo": "452781",
        Equipo: "ca26",
        Descripción: "Cambio componente",
        Prioridad: "Alta",
      },
      {
        "Orden de trabajo": "452782",
        Equipo: "CA29",
        Descripción: "Inspección",
        Prioridad: "Media",
      },
    ]);

    expect(result.rows).toHaveLength(2);
    expect(result.rows[0]).toMatchObject({
      otNumber: "452781",
      equipmentCode: "CA26",
      description: "Cambio componente",
    });
    expect(result.rows[0].extra?.Prioridad).toBe("Alta");
    expect(result.detectedColumns.ot).toBe("Orden de trabajo");
    expect(result.detectedColumns.equipment).toBe("Equipo");
  });

  it("detects SAP-like headers, ignores invalid rows, and counts duplicated OT", () => {
    const result = mapExcelRows([
      {
        Nro_OT: "1001",
        "Código de equipo": "EQ-01",
        "Texto breve": "Revisión",
      },
      {
        Nro_OT: "1001",
        "Código de equipo": "EQ-01",
        "Texto breve": "Duplicada",
      },
      { Nro_OT: "", "Código de equipo": "EQ-02", "Texto breve": "Sin OT" },
      {
        Nro_OT: "",
        "Código de equipo": "",
        "Texto breve": "Fila de título o pie",
      },
    ]);

    expect(result.rows).toHaveLength(2);
    expect(result.duplicates).toBe(1);
    expect(result.invalid).toBe(1);
  });

  it("reads a maintenance-plan header below title rows and combines weekday/day columns", () => {
    const rawRows = rowsFromExcelMatrix([
      ["Programa semanal de mantenimiento"],
      ["", "", "", "", "", "", "", "Lun", "Mar"],
      [
        "ÁREA",
        "TAG",
        "DESCRIPCIÓN DE UBICACIÓN TÉCNICA",
        "ORDEN DE TRABAJO",
        "DESCRIPCIÓN DE ORDEN DE TRABAJO",
        "N° OPERACIÓN",
        "DESCRIPCIÓN DE OPERACIÓN",
        21,
        22,
      ],
      [
        "CARGUIO",
        "CARGUIO",
        "EQ_OEX04",
        6818437,
        "Reparación",
        10,
        "Ajuste",
        4,
        "",
      ],
      [
        "CARGUIO",
        "CARGUIO",
        "EQ_OEX04",
        6818437,
        "Reparación",
        20,
        "Prueba",
        "",
        6,
      ],
      ["Leyenda", "", "", "", "", "", "", "", ""],
    ]);

    expect(rawRows).not.toBeNull();
    expect(rawRows?.[0]).toMatchObject({ "Lun 21": 4, "Mar 22": "" });
    expect(rawRows?.[1]).toMatchObject({ "Lun 21": "", "Mar 22": 6 });

    const result = mapExcelRows(rawRows ?? []);
    expect(result.rows).toHaveLength(1);
    expect(result.rows[0]).toMatchObject({
      otNumber: "6818437",
      equipmentCode: "EX04",
      description: "Reparación",
    });
    expect(result.rows[0].extra?.["N° OPERACIÓN"]).toBe('["10","20"]');
    expect(result.rows[0].extra?.["DESCRIPCIÓN DE OPERACIÓN"]).toBe(
      '["Ajuste","Prueba"]'
    );
    expect(result.rows[0].extra?.["Lun 21"]).toBe("4");
    expect(result.rows[0].extra?.["Mar 22"]).toBe("6");
    expect(result.duplicates).toBe(1);
    expect(result.invalid).toBe(0);
    expect(result.detectedColumns.equipment).toBe(
      "DESCRIPCIÓN DE UBICACIÓN TÉCNICA"
    );
  });

  it("removes the EQ_O prefix from technical-location equipment codes", () => {
    const result = mapExcelRows([
      {
        "ORDEN DE TRABAJO": "1",
        "DESCRIPCIÓN DE UBICACIÓN TÉCNICA": "EQ_OCR02",
      },
      {
        "ORDEN DE TRABAJO": "2",
        "DESCRIPCIÓN DE UBICACIÓN TÉCNICA": "EQ_OCA13",
      },
    ]);

    expect(result.rows.map(row => row.equipmentCode)).toEqual(["CR02", "CA13"]);
  });

  it("uses an explicit bucket for empty technical locations", () => {
    const result = mapExcelRows([
      {
        "ORDEN DE TRABAJO": "3",
        "DESCRIPCIÓN DE UBICACIÓN TÉCNICA": 0,
      },
    ]);

    expect(result.rows).toHaveLength(0);
    expect(result.invalidRows).toEqual([
      { rowNumber: 2, otNumber: "3", reason: "Falta código de Equipo." },
    ]);
  });

  it("reports the Excel row and reason while continuing with valid rows", () => {
    const result = mapExcelRows([
      { OT: "7001", Equipo: "CA01", Descripción: "Válida" },
      { OT: "7002", Equipo: "", Descripción: "Sin equipo" },
      { OT: "7003", Equipo: "CR02", Descripción: "Válida" },
    ]);

    expect(result.rows.map(row => row.otNumber)).toEqual(["7001", "7003"]);
    expect(result.invalidRows).toEqual([
      { rowNumber: 3, otNumber: "7002", reason: "Falta código de Equipo." },
    ]);
  });

  it("fails with a clear message when OT or equipment columns are missing", () => {
    expect(() =>
      mapExcelRows([{ Semana: "SEM 38", Tarea: "Inspección" }])
    ).toThrow("columnas de OT y equipo");
  });
});
