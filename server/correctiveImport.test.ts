import { describe, expect, it } from "vitest";
import { parseCorrectiveLines } from "@shared/correctiveImport";

const currentCrEquipmentCodes = ["CR01", "CR02", "CR05", "CR06"] as const;

describe("Corrective OT parser", () => {
  it.each(currentCrEquipmentCodes)(
    "conserva CR en la descripción y detecta %s como equipo",
    equipmentCode => {
      const result = parseCorrectiveLines(
        `OT 7${equipmentCode.slice(2)}12345 R 4${equipmentCode.slice(2)}12345 CR CAMBIO DE FILTRO ${equipmentCode}`
      );

      expect(result.invalid).toEqual([]);
      expect(result.rows).toHaveLength(1);
      expect(result.rows[0]).toMatchObject({
        equipmentCode,
        reservation: `4${equipmentCode.slice(2)}12345`,
        description: "CR CAMBIO DE FILTRO",
      });
      expect(result.rows[0]?.description).toContain("CR");
    }
  );

  it("detecta OT, reserva, descripción y CR02 aunque cambie el orden", () => {
    const result = parseCorrectiveLines(
      "CR02 CR CAMBIO DE FILTRO OT 6831949 R 3424500"
    );

    expect(result.invalid).toEqual([]);
    expect(result.rows[0]).toEqual({
      otNumber: "6831949",
      reservation: "3424500",
      equipmentCode: "CR02",
      description: "CR CAMBIO DE FILTRO",
    });
  });

  it("crea filas múltiples sin alterar sus códigos CR", () => {
    const result = parseCorrectiveLines(
      currentCrEquipmentCodes
        .map(
          (equipmentCode, index) =>
            `OT 68${String(31949 + index)} R 34${String(24500 + index)} CR CAMBIO CORRECTIVO ${equipmentCode}`
        )
        .join("\n")
    );

    expect(result.invalid).toEqual([]);
    expect(result.rows.map(row => row.equipmentCode)).toEqual([...currentCrEquipmentCodes]);
    expect(result.rows.every(row => row.description.startsWith("CR "))).toBe(true);
  });
});
