import { appendDamagePartOverrides, mapDamagePartFormulas, validateFormula } from "../utils/formulaUtils.js";

export const METAMAGIC_DEFINITION = {
  key: "intensifiedSpell",
  name: "Intensified Spell",
  get prefix() { return globalThis.game?.i18n?.localize?.("NAS.metamagic.prefixes.intensifiedSpell") ?? "Intensified"; },
};

function getCapUpdateFromCount(countExpr, cl) {
  if (!countExpr || !Number.isFinite(cl)) return null;
  const expr = countExpr.replace(/^\(|\)$/g, "");
  const clValue = Number(cl);
  const half = Math.floor(clValue / 2);
  const doubled = clValue * 2;

  const clampMatch = expr.match(/clamp\(\s*floor\(\s*@cl\s*\/\s*2\s*\)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)/i);
  if (clampMatch) {
    const cap = Number(clampMatch[2]);
    if (!Number.isFinite(cap) || half <= cap) return null;
    const nextCap = Math.min(cap + 5, half);
    return {
      updatedExpr: expr.replace(clampMatch[2], String(nextCap)),
      changed: true,
    };
  }

  const minCapHalf = expr.match(/min\(\s*(\d+)\s*,\s*floor\(\s*@cl\s*\/\s*2\s*\)\s*\)/i);
  if (minCapHalf) {
    const cap = Number(minCapHalf[1]);
    if (!Number.isFinite(cap) || half <= cap) return null;
    const nextCap = Math.min(cap + 5, half);
    return {
      updatedExpr: expr.replace(minCapHalf[1], String(nextCap)),
      changed: true,
    };
  }

  const minHalfCap = expr.match(/min\(\s*floor\(\s*@cl\s*\/\s*2\s*\)\s*,\s*(\d+)\s*\)/i);
  if (minHalfCap) {
    const cap = Number(minHalfCap[1]);
    if (!Number.isFinite(cap) || half <= cap) return null;
    const nextCap = Math.min(cap + 5, half);
    return {
      updatedExpr: expr.replace(minHalfCap[1], String(nextCap)),
      changed: true,
    };
  }

  const minCapCl = expr.match(/min\(\s*(\d+)\s*,\s*@cl\s*\)/i);
  if (minCapCl) {
    const cap = Number(minCapCl[1]);
    if (!Number.isFinite(cap) || clValue <= cap) return null;
    const nextCap = Math.min(cap + 5, clValue);
    return {
      updatedExpr: expr.replace(minCapCl[1], String(nextCap)),
      changed: true,
    };
  }

  const minClCap = expr.match(/min\(\s*@cl\s*,\s*(\d+)\s*\)/i);
  if (minClCap) {
    const cap = Number(minClCap[1]);
    if (!Number.isFinite(cap) || clValue <= cap) return null;
    const nextCap = Math.min(cap + 5, clValue);
    return {
      updatedExpr: expr.replace(minClCap[1], String(nextCap)),
      changed: true,
    };
  }

  const minCapDouble = expr.match(/min\(\s*(\d+)\s*,\s*@cl\s*\*\s*2\s*\)/i);
  if (minCapDouble) {
    const cap = Number(minCapDouble[1]);
    if (!Number.isFinite(cap) || doubled <= cap) return null;
    const nextCap = Math.min(cap + 5, doubled);
    return {
      updatedExpr: expr.replace(minCapDouble[1], String(nextCap)),
      changed: true,
    };
  }

  return null;
}

function updateFormula(formula, cl) {
  if (!formula || typeof formula !== "string") return { formula, changed: false };
  const dicePattern = /(\([^)]*\)|@cl|clamp\([^)]*\)|min\([^)]*\)|floor\([^)]*\)|\d+)(\s*\)*)\s*d\s*\d+/gi;
  let match;
  let lastIndex = 0;
  let updated = "";
  let changed = false;
  while ((match = dicePattern.exec(formula)) !== null) {
    const full = match[0];
    const countExpr = match[1];
    const trailingParens = match[2] ?? "";
    const combinedExpr = `${countExpr}${trailingParens}`.trim();
    const replacementInfo = getCapUpdateFromCount(combinedExpr, cl);
    const updatedCountExpr = replacementInfo?.updatedExpr;
    updated += formula.slice(lastIndex, match.index);
    if (updatedCountExpr) {
      const normalizedCount = combinedExpr.startsWith("(") && combinedExpr.endsWith(")")
        ? `(${updatedCountExpr})`
        : updatedCountExpr;
      updated += full.replace(combinedExpr, normalizedCount);
      changed = true;
    } else {
      updated += full;
    }
    lastIndex = match.index + full.length;
  }
  updated += formula.slice(lastIndex);
  return { formula: updated, changed };
}

export function canIntensifyDamageFormula(formula, cl) {
  if (!formula || typeof formula !== "string") return false;
  const clValue = Number(cl ?? 0);
  if (!Number.isFinite(clValue) || clValue <= 0) return false;
  return updateFormula(formula, clValue).changed === true;
}

export function canIntensifyAnyDamagePart(parts, cl) {
  if (!Array.isArray(parts) || !parts.length) return false;
  return parts.some((part) => {
    const formula = Array.isArray(part)
      ? part[0]
      : (part && typeof part === "object" ? part.formula : "");
    return canIntensifyDamageFormula(formula, cl);
  });
}

function intensifyFormula(formula, cl, rollData = {}) {
  const updated = updateFormula(formula, cl);
  if (!updated.changed || updated.formula === formula) return formula;
  if (!validateFormula(updated.formula, rollData)) return formula;
  return updated.formula;
}

export function applyIntensifyToFormula(formula, cl, rollData = {}) {
  return intensifyFormula(formula, cl, rollData);
}

async function evaluateCountExpression(expr, rollData = {}, cl = 0) {
  try {
    const data = { ...(rollData ?? {}), cl };
    const roll = await new Roll(expr, data).evaluate({ async: true });
    const total = Number(roll?.total);
    return Number.isFinite(total) ? total : null;
  } catch (_error) {
    return null;
  }
}

async function applyIntensifyFallbackFromSource(formula, cl, rollData = {}, sourceFormula = "") {
  if (!sourceFormula || typeof sourceFormula !== "string") return formula;

  const sourceDicePattern = /(\([^)]*\)|@cl|clamp\([^)]*\)|min\([^)]*\)|floor\([^)]*\)|\d+)(\s*\)*)\s*d\s*(\d+)/gi;
  /** @type {{faces:number, delta:number, intensifiedCount:number}[]} */
  const deltas = [];
  let match;
  while ((match = sourceDicePattern.exec(sourceFormula)) !== null) {
    const countExpr = match[1];
    const trailingParens = match[2] ?? "";
    const faces = Number(match[3]);
    if (!Number.isFinite(faces) || faces <= 0) continue;
    const combinedExpr = `${countExpr}${trailingParens}`.trim();
    const replacementInfo = getCapUpdateFromCount(combinedExpr, cl);
    if (!replacementInfo?.updatedExpr) continue;
    const baseCount = await evaluateCountExpression(combinedExpr, rollData, cl);
    const intensifiedCount = await evaluateCountExpression(replacementInfo.updatedExpr, rollData, cl);
    if (!Number.isFinite(baseCount) || !Number.isFinite(intensifiedCount)) continue;
    const delta = intensifiedCount - baseCount;
    if (delta > 0) deltas.push({ faces, delta, intensifiedCount });
  }

  if (!deltas.length) return formula;

  let nextFormula = formula;
  const replaceExpressionCountBeforeDie = (inputFormula, faces, intensifiedCount) => {
    const dieRegex = new RegExp(`d\\s*${faces}\\b`, "i");
    const dieMatch = dieRegex.exec(inputFormula);
    if (!dieMatch) return inputFormula;

    const dieIndex = dieMatch.index;
    let i = dieIndex - 1;
    while (i >= 0 && /\s/.test(inputFormula[i])) i -= 1;
    if (i < 0) return inputFormula;

    let start = i;
    if (inputFormula[i] === ")") {
      let depth = 0;
      while (i >= 0) {
        const ch = inputFormula[i];
        if (ch === ")") depth += 1;
        if (ch === "(") {
          depth -= 1;
          if (depth === 0) {
            start = i;
            break;
          }
        }
        i -= 1;
      }
      if (depth !== 0) return inputFormula;
    } else {
      while (i >= 0 && /[A-Za-z0-9_@.]/.test(inputFormula[i])) i -= 1;
      start = i + 1;
    }

    if (start >= dieIndex) return inputFormula;
    const countExpr = inputFormula.slice(start, dieIndex).trim();
    if (!countExpr) return inputFormula;

    return `${inputFormula.slice(0, start)}${intensifiedCount}${inputFormula.slice(dieIndex)}`;
  };

  for (const { faces, delta, intensifiedCount } of deltas) {
    const faceRegex = new RegExp(`\\b(\\d+)\\s*d\\s*${faces}\\b`);
    const faceMatch = nextFormula.match(faceRegex);
    if (faceMatch) {
      const baseCount = Number(faceMatch[1]);
      if (!Number.isFinite(baseCount)) continue;
      const nextCount = baseCount + delta;
      nextFormula = nextFormula.replace(faceRegex, `${nextCount}d${faces}`);
      continue;
    }

    if (Number.isFinite(intensifiedCount) && intensifiedCount > 0) {
      nextFormula = replaceExpressionCountBeforeDie(nextFormula, faces, intensifiedCount);
    }
  }

  return nextFormula;
}

export async function applyIntensifyToFormulaLate(formula, cl, rollData = {}, { sourceFormula } = {}) {
  const direct = intensifyFormula(formula, cl, rollData);
  if (direct !== formula) return direct;
  return applyIntensifyFallbackFromSource(formula, cl, rollData, sourceFormula);
}

export function applyIntensifiedSpell(context, cl, rollData = {}) {
  const clValue = Number(cl ?? context?.actor?.system?.attributes?.casterLevel?.value ?? 0);
  if (!Number.isFinite(clValue) || clValue <= 0) return false;

  const overrides = mapDamagePartFormulas(
    context,
    (formula) => intensifyFormula(formula, clValue, rollData)
  );
  if (!appendDamagePartOverrides(context, overrides)) return false;

  if (!context.metamagic) {
    context.metamagic = { applied: [], slotIncrease: 0 };
  }

  if (!context.metamagic.applied.includes(METAMAGIC_DEFINITION.name)) {
    context.metamagic.applied.push(METAMAGIC_DEFINITION.name);
    context.metamagic.slotIncrease += 1;
  }

  return true;
}
