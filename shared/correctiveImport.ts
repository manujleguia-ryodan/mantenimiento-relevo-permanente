export type CorrectiveDraft = {
  otNumber: string;
  equipmentCode: string;
  reservation: string;
  description: string;
};

export function parseCorrectiveLines(value: string) {
  const rows: CorrectiveDraft[] = [];
  const invalid: string[] = [];

  for (const rawLine of value
    .split(/\r?\n/)
    .map(line => line.trim())
    .filter(Boolean)) {
    const tokens = Array.from(rawLine.matchAll(/\b[A-Z0-9_-]+\b/gi)).map(match => ({
      value: match[0],
      index: match.index ?? 0,
    }));
    const otToken = tokens.find(token => /^(?:OT)?[67]\d{5,}$/i.test(token.value));
    const reservationToken = tokens.find(token => /^(?:R)?[34]\d{5,}$/i.test(token.value));
    const equipmentToken = tokens.find(
      token =>
        /^[A-Z]{1,6}\d{1,4}$/i.test(token.value) &&
        !/^(?:OT|R)\d+$/i.test(token.value)
    );

    if (!otToken || !equipmentToken) {
      invalid.push(rawLine);
      continue;
    }

    const excluded = new Set([otToken.index, equipmentToken.index]);
    if (reservationToken) excluded.add(reservationToken.index);
    const description = tokens
      .filter(token => !excluded.has(token.index) && !/^(?:OT|R)$/i.test(token.value))
      .map(token => token.value)
      .join(" ")
      .trim();

    if (!description) {
      invalid.push(rawLine);
      continue;
    }

    rows.push({
      otNumber: otToken.value.replace(/^OT/i, ""),
      reservation: reservationToken?.value.replace(/^R/i, "") ?? "",
      equipmentCode: equipmentToken.value.toUpperCase(),
      description,
    });
  }

  return { rows, invalid };
}
