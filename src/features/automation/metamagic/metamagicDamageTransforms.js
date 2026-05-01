import { applyEmpowerToFormula } from "./empowerSpell.js";
import { applyIntensifyToFormulaLate } from "./intensifiedSpell.js";
import { applyMaximizeToFormula } from "./maximizeSpell.js";

export async function applyMetamagicDamageTransforms({
  formula,
  transforms = [],
  rollData = {},
  maximizeAlreadyApplied = false,
  intensifySourceFormula = null
} = {}) {
  let nextFormula = (formula ?? "").toString();
  const applied = [];
  for (const transformKey of transforms) {
    if (transformKey === "intensifiedSpell") {
      const cl = Number(rollData?.cl ?? 0);
      nextFormula = await applyIntensifyToFormulaLate(nextFormula, cl, rollData, {
        sourceFormula: intensifySourceFormula
      });
      applied.push(transformKey);
      continue;
    }
    if (transformKey === "maximizeSpell") {
      if (maximizeAlreadyApplied) continue;
      nextFormula = applyMaximizeToFormula(nextFormula, rollData);
      applied.push(transformKey);
      continue;
    }
    if (transformKey === "empowerSpell") {
      nextFormula = applyEmpowerToFormula(nextFormula);
      applied.push(transformKey);
      continue;
    }
  }

  return {
    formula: nextFormula,
    applied
  };
}
