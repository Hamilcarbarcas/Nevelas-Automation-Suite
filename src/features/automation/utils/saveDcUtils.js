const NO_SAVE_TYPE_TOKENS = new Set(["", "none", "null", "—", "-"]);

function normalizeSaveTypeToken(raw) {
  return (raw ?? "").toString().trim().toLowerCase();
}

/**
 * True when the spell/action exposes a saving throw (fort/ref/will, etc.).
 * PF1 omits or clears `save.type` when the spell does not allow a save.
 */
export function contextHasSavingThrow(context, action) {
  const candidates = [
    context?.save?.type,
    action?.action?.save?.type,
    action?.item?.system?.save?.type,
  ];
  for (const raw of candidates) {
    const t = normalizeSaveTypeToken(raw);
    if (!t || NO_SAVE_TYPE_TOKENS.has(t)) continue;
    return true;
  }
  return false;
}

/**
 * Snapshot pre-modifier DC for chat "Base DC" when no metamagic selections run
 * (early return) or other paths skip the main baseDc assignment.
 */
export function ensureSpellSaveBaseDcSnapshot(action, context) {
  if (action?.item?.type !== "spell") return;
  if (!contextHasSavingThrow(context, action)) return;
  const existing = Number(context?.save?.baseDc);
  if (Number.isFinite(existing)) return;
  const v = resolveFeatSaveDcBase(action, context);
  if (!Number.isFinite(v)) return;
  context.save ??= {};
  context.save.baseDc = v;
  context.metamagic ??= { applied: [], slotIncrease: 0 };
  context.metamagic.baseSaveDc = v;
}

/**
 * Numeric save DC for feat bonuses (Maleficium, Eldritch Researcher, etc.).
 * `context.save.dc` from collectSpellActionData is the raw item field (often a formula
 * or empty). Coercing it with Number() yields 0 or NaN; use evaluated roll + getDC first.
 */
export function resolveFeatSaveDcBase(action, context) {
  const contextDcNum = context?.save?.dc;
  if (typeof contextDcNum === "number" && Number.isFinite(contextDcNum)) {
    return contextDcNum;
  }

  const evalTotal = Number(context?.save?.evaluated?.total);
  if (Number.isFinite(evalTotal) && evalTotal > 0) {
    return evalTotal;
  }

  if (typeof action?.action?.getDC === "function" && action?.shared?.rollData) {
    const liveDc = Number(action.action.getDC(action.shared.rollData));
    if (Number.isFinite(liveDc)) {
      return liveDc;
    }
  }

  const contextDcRaw = context?.save?.dc;
  if (typeof contextDcRaw === "string" && contextDcRaw.trim().length > 0) {
    const parsed = Number(contextDcRaw.trim());
    if (Number.isFinite(parsed)) {
      return parsed;
    }
  }

  const sharedDcRaw = action?.shared?.saveDC;
  if (typeof sharedDcRaw === "number" && Number.isFinite(sharedDcRaw)) {
    return sharedDcRaw;
  }
  if (typeof sharedDcRaw === "string" && sharedDcRaw.trim().length > 0) {
    const parsed = Number(sharedDcRaw.trim());
    if (Number.isFinite(parsed)) {
      return parsed;
    }
  }

  if (Number.isFinite(evalTotal)) {
    return evalTotal;
  }

  return NaN;
}
