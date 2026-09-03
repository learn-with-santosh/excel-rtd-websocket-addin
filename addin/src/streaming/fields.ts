export const VALID_FIELDS = ["LAST", "OPEN", "HIGH", "LOW", "VOLUME"] as const;

export type FieldName = (typeof VALID_FIELDS)[number];

/**
 * Parses "last, open,high" into ["LAST", "OPEN", "HIGH"].
 * Throws a plain Error with a user-friendly message — the caller
 * converts it into a CustomFunctions error.
 */
export function parseFields(raw: string): FieldName[] {
  const tokens = (raw ?? "")
    .split(",")
    .map((t) => t.trim().toUpperCase())
    .filter((t) => t.length > 0);

  if (tokens.length === 0) {
    throw new Error(
      'Provide at least one field, e.g. "LAST" or "LAST,OPEN,HIGH,LOW,VOLUME".'
    );
  }

  const picked: FieldName[] = [];
  for (const token of tokens) {
    if ((VALID_FIELDS as readonly string[]).includes(token)) {
      if (!picked.includes(token as FieldName)) picked.push(token as FieldName);
    } else {
      const hint = token === "VOLUMN" ? " Did you mean VOLUME?" : "";
      throw new Error(
        `Unknown field "${token}".${hint} Valid fields: ${VALID_FIELDS.join(", ")}`
      );
    }
  }
  return picked;
}