function getRollImplementation() {
  return Roll?.defaultImplementation ?? Roll;
}

function toFiniteInt(value) {
  const n = Number(value);
  if (!Number.isFinite(n)) return 0;
  return Math.max(0, Math.floor(n));
}

function normalizeFormula(formula) {
  return (formula ?? "").toString().trim();
}

function tryEvaluateRoll(roll) {
  if (!roll) return null;
  try {
    if (!roll.evaluated && typeof roll.evaluateSync === "function") {
      roll.evaluateSync({ minimize: true, strict: false, allowStrings: true });
    }
    return roll;
  } catch (_err) {
    return null;
  }
}

function buildRoll(formula, rollData = {}) {
  const RollImpl = getRollImplementation();
  const raw = normalizeFormula(formula);
  if (!RollImpl || !raw) return null;
  try {
    return new RollImpl(raw, rollData);
  } catch (_err) {
    return null;
  }
}

function getDieCount(dieTerm) {
  if (!dieTerm) return 0;
  const direct = toFiniteInt(dieTerm.number);
  if (direct > 0) return direct;

  const numberRoll = dieTerm._number;
  const evaluated = tryEvaluateRoll(numberRoll);
  if (!evaluated) return 0;
  return toFiniteInt(evaluated.total);
}

function getCurrentFormulaOverride(context, index, fallback) {
  const overrides = context?.damageOverrides?.parts;
  if (Array.isArray(overrides)) {
    for (let i = overrides.length - 1; i >= 0; i -= 1) {
      const entry = overrides[i];
      if (entry?.index === index && entry?.formula) return entry.formula;
    }
  }
  return fallback;
}

function getDiceCountExpression(dieTerm) {
  const number = dieTerm?.number;
  if (Number.isFinite(number)) return String(toFiniteInt(number));

  const expr =
    dieTerm?._number?.formula
    ?? dieTerm?.number?.formula
    ?? "";
  const raw = normalizeFormula(expr);
  if (!raw) return "0";

  const simple = /^[a-z0-9_@.]+$/i.test(raw);
  return simple ? raw : `(${raw})`;
}

export function validateFormula(formula, rollData = {}) {
  const roll = buildRoll(formula, rollData);
  if (!roll) return false;
  return Boolean(tryEvaluateRoll(roll));
}

export function mapDamagePartFormulas(context, mapper) {
  if (!context?.damage?.parts || !Array.isArray(context.damage.parts)) return [];
  if (typeof mapper !== "function") return [];

  const overrides = [];
  context.damage.parts.forEach((part, index) => {
    if (!part || typeof part !== "object") return;
    const baseFormula = part.formula ?? part[0];
    const formula = getCurrentFormulaOverride(context, index, baseFormula);
    const raw = normalizeFormula(formula);
    if (!raw) return;

    const next = mapper(raw, {
      index,
      part,
      isArray: Array.isArray(part),
      baseFormula
    });
    const nextFormula = normalizeFormula(next);
    if (!nextFormula || nextFormula === raw) return;

    overrides.push({
      index,
      isArray: Array.isArray(part),
      formula: nextFormula
    });
  });
  return overrides;
}

export function appendDamagePartOverrides(context, overrides) {
  if (!Array.isArray(overrides) || !overrides.length) return false;
  context.damageOverrides ??= { parts: [] };
  context.damageOverrides.parts = [
    ...(context.damageOverrides.parts ?? []),
    ...overrides
  ];
  return true;
}

export function maximizeDiceTermsEngine(formula, rollData = {}) {
  const roll = buildRoll(formula, rollData);
  if (!roll) return null;
  const terms = Array.isArray(roll.terms) ? roll.terms : [];
  if (!terms.length) return null;

  let changed = false;
  const out = terms.map((term) => {
    if (!(term instanceof foundry.dice.terms.DiceTerm)) {
      return term?.formula ?? "";
    }
    changed = true;
    const countExpr = getDiceCountExpression(term);
    const faces = toFiniteInt(term?.faces);
    return `(${countExpr} * ${faces})`;
  }).join("");

  if (!changed) return null;
  if (!validateFormula(out, rollData)) return null;
  return out;
}

export function getDiceCountFromRoll(roll) {
  const evaluated = tryEvaluateRoll(roll);
  if (!evaluated) return 0;
  const dice = Array.isArray(evaluated.dice) ? evaluated.dice : [];
  if (!dice.length) return 0;
  return dice.reduce((sum, die) => sum + getDieCount(die), 0);
}

export function getDiceCountFromFormula(formula, rollData = {}) {
  const roll = buildRoll(formula, rollData);
  if (!roll) return 0;
  return getDiceCountFromRoll(roll);
}

function evaluateFormulaAtCl(formula, rollData, cl) {
  const baseData = rollData && typeof rollData === "object" ? rollData : {};
  const data = { ...baseData, cl };
  const roll = buildRoll(formula, data);
  if (!roll) return null;
  const evaluated = tryEvaluateRoll(roll);
  if (!evaluated) return null;
  const total = Number(evaluated.total);
  return Number.isFinite(total) ? total : null;
}

function looksCasterLevelScaled(formula) {
  if (typeof formula !== "string") return false;
  return /@cl|\/\s*level|caster\s*level|per\s*level/i.test(formula);
}

export function getCasterLevelEquivalentFromFormula(formula, { rollData = {}, casterLevel = 0 } = {}) {
  const cl = toFiniteInt(casterLevel);
  if (!cl) return 0;
  if (!looksCasterLevelScaled(formula)) return 0;

  const atCurrent = evaluateFormulaAtCl(formula, rollData, cl);
  if (!Number.isFinite(atCurrent)) return 0;

  const at1 = evaluateFormulaAtCl(formula, rollData, 1);
  const at2 = evaluateFormulaAtCl(formula, rollData, 2);
  const step = Number.isFinite(at1) && Number.isFinite(at2) ? (at2 - at1) : 0;

  if (step > 0) {
    return toFiniteInt(atCurrent / step);
  }

  return cl;
}

export function getWoundMagnitudeFromFormula({
  formula,
  rollData = {},
  casterLevel = 0,
  fallback = 0
} = {}) {
  const dice = getDiceCountFromFormula(formula, rollData);
  if (dice > 0) return dice;

  const clEquivalent = getCasterLevelEquivalentFromFormula(formula, { rollData, casterLevel });
  if (clEquivalent > 0) return clEquivalent;

  return toFiniteInt(Math.abs(Number(fallback) || 0));
}
