export const METAMAGIC_DEFINITION = {
  key: "enlargeSpell",
  name: "Enlarge Spell",
  get prefix() { return globalThis.game?.i18n?.localize?.("NAS.metamagic.prefixes.enlargeSpell") ?? "Enlarge"; },
};

const ENLARGE_RANGE_UNITS = new Set(["close", "medium", "long"]);

function getNumeric(value) {
  const num = Number(value);
  return Number.isFinite(num) ? num : null;
}

function resolveBaseRangeFromRollData(units, action) {
  if (!units || !action) return null;
  const rollData = action?.shared?.rollData ?? null;
  try {
    const total = pf1?.utils?.calculateRangeFormula?.("", units, rollData ?? {});
    const num = Number(total);
    return Number.isFinite(num) ? num : null;
  } catch (_err) {
    return null;
  }
}

function pickPositiveRange(...values) {
  for (const value of values) {
    const num = getNumeric(value);
    if (Number.isFinite(num) && num > 0) return num;
  }
  return null;
}

export function canApplyEnlargeSpell(context) {
  const units = (context?.range?.range?.units ?? "").toString().trim().toLowerCase();
  return ENLARGE_RANGE_UNITS.has(units);
}

export function applyEnlargeSpell(context, action = null) {
  const units = (context?.range?.range?.units ?? "").toString().trim().toLowerCase();
  const evaluatedTotalRaw = context?.range?.evaluated?.value?.total ?? null;
  const rangeValueRaw = context?.range?.range?.value ?? null;
  const computedFromRollData = resolveBaseRangeFromRollData(units, action);
  const baseRange = pickPositiveRange(evaluatedTotalRaw, rangeValueRaw, computedFromRollData);
  const canApply = canApplyEnlargeSpell(context);
  if (!context?.range?.range) return false;
  if (!canApply) return false;

  if (!Number.isFinite(baseRange) || baseRange <= 0) return false;

  const nextRange = baseRange * 2;
  context.range.range.units = "ft";
  context.range.range.value = String(nextRange);
  context.range.hasRange = true;
  context.range.isRanged = true;
  context.range.touch = false;

  const baseMin =
    getNumeric(context?.range?.evaluated?.minValue?.total)
    ?? getNumeric(context?.range?.range?.minValue);
  if (Number.isFinite(baseMin) && baseMin >= 0) {
    context.range.range.minUnits = "ft";
    context.range.range.minValue = String(baseMin * 2);
  }

  context.metamagic ??= { applied: [], slotIncrease: 0 };
  if (!context.metamagic.applied.includes(METAMAGIC_DEFINITION.name)) {
    context.metamagic.applied.push(METAMAGIC_DEFINITION.name);
    context.metamagic.slotIncrease = Number(context.metamagic.slotIncrease ?? 0) + 1;
  }

  return true;
}
