export const METAMAGIC_DEFINITION = {
  key: "extendSpell",
  name: "Extend Spell",
  get prefix() { return globalThis.game?.i18n?.localize?.("NAS.metamagic.prefixes.extendSpell") ?? "Extend"; },
};

const PF1_DURATION_UNITS_WITHOUT_VALUE_INPUT = new Set(["", "turn", "inst", "perm", "seeText"]);

function durationUnitAllowsValueInput(units) {
  return !PF1_DURATION_UNITS_WITHOUT_VALUE_INPUT.has((units ?? "").toString().toLowerCase());
}

function durationValueIndicatesNonExtendableSemantics(rawValue) {
  const s = (rawValue ?? "").toString().trim().toLowerCase();
  if (!s.length) return false;
  if (/^instantaneous\b/.test(s) || /^instant\b/.test(s) || s === "instant" || s === "inst") return true;
  if (/^permanent\b/.test(s) || /^perm\b/.test(s)) return true;
  if (/^concentration\b/.test(s)) return true;
  return false;
}

export function isDurationEligibleForExtendSpell(durationLike) {
  if (!durationLike || durationLike.concentration) return false;
  const units = (durationLike.units ?? "").toString().toLowerCase();
  if (units === "inst" || units === "instantaneous" || units === "perm" || units === "permanent") {
    return false;
  }
  if (durationUnitAllowsValueInput(units) && durationValueIndicatesNonExtendableSemantics(durationLike.value)) {
    return false;
  }
  return true;
}

function applyExtendDurationOnlyInternal(context) {
  if (!context?.duration) return false;
  if (context._nasExtendDurationApplied === true) return false;

  if (!isDurationEligibleForExtendSpell(context.duration)) return false;

  const baseTotal = Number(context.duration.evaluated?.total ?? 0);
  const extendedTotal = baseTotal * 2;
  context.duration.value = Number.isFinite(extendedTotal) ? String(extendedTotal) : context.duration.value;
  context.duration.evaluated = {
    ...(context.duration.evaluated ?? {}),
    total: Number.isFinite(extendedTotal) ? extendedTotal : baseTotal,
  };
  context._nasExtendDurationApplied = true;
  return true;
}

export function applyExtendDurationOnly(context) {
  return applyExtendDurationOnlyInternal(context);
}

export function applyExtendSpell(context) {
  const didApplyDuration = applyExtendDurationOnlyInternal(context);
  if (!didApplyDuration) return false;

  if (!context.metamagic) {
    context.metamagic = { applied: [], slotIncrease: 0 };
  }

  if (!context.metamagic.applied.includes(METAMAGIC_DEFINITION.name)) {
    context.metamagic.applied.push(METAMAGIC_DEFINITION.name);
    context.metamagic.slotIncrease += 1;
  }

  return didApplyDuration;
}

export function applyExtendSpellWithMaskFocus(context) {
  if (!context?.duration) return false;
  if (context._nasExtendDurationApplied === true) return false;

  if (!isDurationEligibleForExtendSpell(context.duration)) return false;

  const baseTotal = Number(context.duration.evaluated?.total ?? 0);
  if (!Number.isFinite(baseTotal) || baseTotal <= 0) {
    return false;
  }

  const extendedSelfTotal = baseTotal * 2;
  context.duration.maskFocusSelf = {
    baseTotal,
    extendedSelfTotal,
    units: context.duration.units ?? ""
  };
  context.duration.value = String(baseTotal);
  context.duration.evaluated = {
    ...(context.duration.evaluated ?? {}),
    total: baseTotal
  };
  context._nasExtendDurationApplied = true;

  context.metamagic ??= { applied: [], slotIncrease: 0 };
  if (!context.metamagic.applied.includes(METAMAGIC_DEFINITION.name)) {
    context.metamagic.applied.push(METAMAGIC_DEFINITION.name);
  }
  context.metamagic.extendSlotWaivedByMaskFocus = true;

  return true;
}
