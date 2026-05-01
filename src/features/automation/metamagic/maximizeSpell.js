import {
  appendDamagePartOverrides,
  mapDamagePartFormulas,
  maximizeDiceTermsEngine,
  validateFormula
} from "../utils/formulaUtils.js";

export const METAMAGIC_DEFINITION = {
  key: "maximizeSpell",
  name: "Maximize Spell",
  get prefix() { return globalThis.game?.i18n?.localize?.("NAS.metamagic.prefixes.maximizeSpell") ?? "Maximize"; },
};

function updateFormulaWithRegex(formula) {
  if (!formula || typeof formula !== "string") return { formula, changed: false };
  const dicePattern =
    /(\([^)]*\)|@cl|clamp\([^)]*\)|min\([^)]*\)|floor\([^)]*\)|\d+)(\s*\)*)\s*d\s*(\d+)/gi;
  let match;
  let lastIndex = 0;
  let updated = "";
  let changed = false;
  while ((match = dicePattern.exec(formula)) !== null) {
    const full = match[0];
    const countExpr = match[1];
    const trailingParens = match[2] ?? "";
    const faces = match[3];
    const combinedExpr = `${countExpr}${trailingParens}`.trim();
    const replacement = `(${combinedExpr} * ${faces})`;
    updated += formula.slice(lastIndex, match.index);
    updated += replacement;
    changed = true;
    lastIndex = match.index + full.length;
  }
  updated += formula.slice(lastIndex);
  return { formula: updated, changed };
}

function maximizeFormula(formula, rollData = {}) {
  const engineFormula = maximizeDiceTermsEngine(formula, rollData);
  if (engineFormula && engineFormula !== formula) return engineFormula;

  const updated = updateFormulaWithRegex(formula);
  if (!updated.changed || updated.formula === formula) return formula;
  if (!validateFormula(updated.formula, rollData)) return formula;
  return updated.formula;
}

export function applyMaximizeToFormula(formula, rollData = {}) {
  return maximizeFormula(formula, rollData);
}

export function applyMaximizeSpell(context, rollData = {}) {
  const overrides = mapDamagePartFormulas(context, (formula) => maximizeFormula(formula, rollData));
  if (!appendDamagePartOverrides(context, overrides)) return false;

  if (!context.metamagic) {
    context.metamagic = { applied: [], slotIncrease: 0 };
  }

  if (!context.metamagic.applied.includes(METAMAGIC_DEFINITION.name)) {
    context.metamagic.applied.push(METAMAGIC_DEFINITION.name);
    context.metamagic.slotIncrease += 3;
  }

  return true;
}
