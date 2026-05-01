function getDictionary(item) {
  const dictionary = item?.system?.flags?.dictionary;
  return dictionary && typeof dictionary === "object" ? dictionary : {};
}

function normalizeString(value) {
  return (value ?? "").toString().trim();
}

export function getDictionaryString(item, key, { normalize = false } = {}) {
  const dictionary = getDictionary(item);
  const value = normalizeString(dictionary?.[key]);
  return normalize ? value.toLowerCase() : value;
}

export function getDictionaryBoolean(item, key, defaultValue = false) {
  const dictionary = getDictionary(item);
  const raw = dictionary?.[key];
  if (typeof raw === "boolean") return raw;
  if (typeof raw === "number") return raw !== 0;
  if (typeof raw === "string") {
    const normalized = raw.trim().toLowerCase();
    if (normalized === "true" || normalized === "1" || normalized === "yes" || normalized === "on") return true;
    if (normalized === "false" || normalized === "0" || normalized === "no" || normalized === "off") return false;
  }
  return defaultValue;
}

export function getDictionaryNumber(item, key, defaultValue = 0) {
  const dictionary = getDictionary(item);
  const raw = Number(dictionary?.[key]);
  return Number.isFinite(raw) ? raw : defaultValue;
}

export function getDictionaryPrefixedStrings(item, prefix) {
  const dictionary = getDictionary(item);
  return Object.entries(dictionary)
    .filter(([key]) => key.startsWith(prefix))
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([, value]) => normalizeString(value))
    .filter(Boolean);
}

export async function setDictionaryEntries(item, entries = {}, { removeKeys = [] } = {}) {
  if (!item || typeof item?.update !== "function") return;
  const updateData = {};
  Object.entries(entries ?? {}).forEach(([key, value]) => {
    updateData[`system.flags.dictionary.${key}`] = value;
  });
  (Array.isArray(removeKeys) ? removeKeys : []).forEach((key) => {
    updateData[`system.flags.dictionary.-=${key}`] = null;
  });
  if (!Object.keys(updateData).length) return;
  await item.update(updateData);
}

export async function setDictionaryPrefixedStrings(item, prefix, values = [], { countKey = null } = {}) {
  const dictionary = getDictionary(item);
  const normalizedValues = Array.from(
    new Set(
      (Array.isArray(values) ? values : [])
        .map((value) => normalizeString(value))
        .filter(Boolean)
    )
  );
  const keepKeys = new Set(normalizedValues.map((_, index) => `${prefix}${index + 1}`));
  const removeKeys = Object.keys(dictionary)
    .filter((key) => key.startsWith(prefix) && !keepKeys.has(key));
  const entries = {};
  normalizedValues.forEach((value, index) => {
    entries[`${prefix}${index + 1}`] = value;
  });
  if (countKey) {
    entries[countKey] = String(normalizedValues.length);
  }
  await setDictionaryEntries(item, entries, { removeKeys });
}

export function findChoiceByStoredName(choices, storedName) {
  const selected = normalizeString(storedName).toLowerCase();
  if (!selected) return null;
  return (Array.isArray(choices) ? choices : []).find(
    (choice) => normalizeString(choice?.label).toLowerCase() === selected
  ) ?? null;
}

export function findChoiceByIdentifier(choices, selectedIdentifier) {
  const selected = normalizeString(selectedIdentifier);
  if (!selected) return null;
  const selectedId = selected.split(".").pop();
  return (Array.isArray(choices) ? choices : []).find((choice) => {
    const uuid = normalizeString(choice?.uuid);
    const id = normalizeString(choice?.id);
    if (uuid && uuid === selected) return true;
    if (id && id === selected) return true;
    if (selectedId && id && id === selectedId) return true;
    return false;
  }) ?? null;
}
