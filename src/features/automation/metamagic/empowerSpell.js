import { appendDamagePartOverrides, mapDamagePartFormulas } from "../utils/formulaUtils.js";
import { MODULE } from "../../../common/module.js";
import { elementFromHtmlLike } from "../../../common/foundryCompat.js";

export const METAMAGIC_DEFINITION = {
  key: "empowerSpell",
  name: "Empower Spell",
  get prefix() { return globalThis.game?.i18n?.localize?.("NAS.metamagic.prefixes.empowerSpell") ?? "Empower"; },
};

function wrapEmpowerFormula(formula) {
  const raw = (formula ?? "").toString().trim();
  if (!raw) return raw;
  return `floor((${raw}) * 1.5)`;
}

export function applyEmpowerToFormula(formula) {
  return wrapEmpowerFormula(formula);
}

export function applyEmpowerSpell(context) {
  const overrides = mapDamagePartFormulas(context, (formula) => wrapEmpowerFormula(formula));
  if (!appendDamagePartOverrides(context, overrides)) return false;

  if (!context.metamagic) {
    context.metamagic = { applied: [], slotIncrease: 0 };
  }

  if (!context.metamagic.applied.includes(METAMAGIC_DEFINITION.name)) {
    context.metamagic.applied.push(METAMAGIC_DEFINITION.name);
    context.metamagic.slotIncrease += 2;
  }

  return true;
}

function decodeRollData(raw) {
  if (!raw || typeof raw !== "string") return null;
  try {
    return JSON.parse(decodeURIComponent(raw));
  } catch {
    try {
      return JSON.parse(unescape(raw));
    } catch {
      return null;
    }
  }
}

function isNasMetamagicEmpowerRoll(rollData) {
  const transforms = rollData?.options?.nasLabelContext?.transforms;
  return Array.isArray(transforms) && transforms.includes(METAMAGIC_DEFINITION.key);
}

function findInnerTotal(node) {
  if (!node || typeof node !== "object") return null;
  const formula = node.formula ?? node.term ?? "";
  const hasEmpower = typeof formula === "string" && formula.replace(/\s+/g, "").includes("*1.5");
  if (!hasEmpower && Number.isFinite(node.total)) return node.total;

  if (Array.isArray(node.terms)) {
    for (const term of node.terms) {
      const inner = findInnerTotal(term);
      if (Number.isFinite(inner)) return inner;
    }
  }
  if (Array.isArray(node.rolls)) {
    for (const roll of node.rolls) {
      const inner = findInnerTotal(roll);
      if (Number.isFinite(inner)) return inner;
    }
  }
  if (node.roll && typeof node.roll === "object") {
    const inner = findInnerTotal(node.roll);
    if (Number.isFinite(inner)) return inner;
  }
  return null;
}

function parseNumericText(value) {
  if (value === null || value === undefined) return null;
  const normalized = `${value}`.replace(/[^0-9.+-]/g, "");
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : null;
}

function sumDicePartTotals(wrapper) {
  if (!wrapper) return null;
  let total = 0;
  let foundAny = false;
  const diceParts = wrapper.querySelectorAll(".tooltip-part");
  for (const part of diceParts) {
    const hasDice = Boolean(part.querySelector(".dice-rolls li, .dice-rolls .roll"));
    if (!hasDice) continue;
    const totalEl = part.querySelector(".part-total");
    const value = parseNumericText(totalEl?.textContent ?? "");
    if (!Number.isFinite(value)) continue;
    total += value;
    foundAny = true;
  }
  return foundAny ? total : null;
}

function hasSystemNumericParts(wrapper) {
  if (!wrapper) return false;
  const parts = wrapper.querySelectorAll(".tooltip-part");
  for (const part of parts) {
    if (part.dataset?.nasEmpowerBonus === "true") continue;
    if (part.dataset?.nasBaseBonus === "true") continue;
    const hasDice = Boolean(part.querySelector(".dice-rolls li, .dice-rolls .roll"));
    if (!hasDice) return true;
  }
  return false;
}

function inferBaseBonusLabel(formula) {
  const normalized = (formula ?? "").toString();
  if (/@cl\b/i.test(normalized) || /resolvedcl/i.test(normalized) || /min\s*\(\s*\d+\s*,\s*\d+\s*\)/i.test(normalized)) {
    return game.i18n?.localize?.("PF1.CasterLevel") || "Caster level";
  }
  return game.i18n?.localize?.("PF1.Bonus") || "Bonus";
}

function ensureNumericBonusRow(wrapper, bonusValue, label, rowKey) {
  if (!wrapper) return false;
  const rowAttributeKey = rowKey.replace(/[A-Z]/g, (m) => `-${m.toLowerCase()}`);
  const existing = wrapper.querySelector(`.tooltip-part[data-${rowAttributeKey}="true"]`);
  if (!Number.isFinite(bonusValue) || bonusValue <= 0) {
    if (existing) {
      existing.remove();
      return true;
    }
    return false;
  }

  let part = existing;
  if (!part) {
    part = document.createElement("section");
    part.classList.add("tooltip-part");
    part.dataset[rowKey] = "true";

    const dice = document.createElement("div");
    dice.classList.add("dice");

    const header = document.createElement("header");
    header.classList.add("part-header", "flexrow");

    const formula = document.createElement("span");
    formula.classList.add("part-formula");

    const flavor = document.createElement("span");
    flavor.classList.add("part-flavor");

    const total = document.createElement("span");
    total.classList.add("part-total");

    const prefix = document.createElement("span");
    prefix.classList.add("part-prefix");

    header.appendChild(formula);
    header.appendChild(flavor);
    header.appendChild(total);
    dice.appendChild(header);
    part.appendChild(dice);
    wrapper.appendChild(part);
    total.appendChild(prefix);
    total.appendChild(document.createTextNode("0"));
  }

  const flavorEl = part.querySelector(".part-flavor");
  if (flavorEl) {
    flavorEl.classList.remove("placeholder");
    flavorEl.textContent = label;
  }

  const totalEl = part.querySelector(".part-total");
  if (totalEl) {
    const prefixEl = totalEl.querySelector(".part-prefix");
    if (prefixEl) {
      prefixEl.textContent = "+";
      const trailing = Array.from(totalEl.childNodes).find((n) => n.nodeType === Node.TEXT_NODE);
      if (trailing) trailing.textContent = `${bonusValue}`;
      else totalEl.appendChild(document.createTextNode(`${bonusValue}`));
    } else {
      totalEl.textContent = `+${bonusValue}`;
    }
  }

  return true;
}

function ensureEmpowerBonusRow(wrapper, bonusValue, label) {
  return ensureNumericBonusRow(wrapper, bonusValue, label, "nasEmpowerBonus");
}

function ensureBaseBonusRow(wrapper, bonusValue, label) {
  return ensureNumericBonusRow(wrapper, bonusValue, label, "nasBaseBonus");
}

export function applyEmpowerTooltipOverrides(html) {
  let tooltipMode = "off";
  try {
    tooltipMode = game.settings.get(MODULE.ID, "enhancedDiceTooltipMode") || "off";
  } catch (_err) {}

  const root = elementFromHtmlLike(html);
  if (!root?.querySelectorAll) return;
  const rolls = root.querySelectorAll(".inline-roll[data-roll]");
  rolls.forEach((rollEl) => {
    const rollData = decodeRollData(rollEl.dataset.roll);
    if (!isNasMetamagicEmpowerRoll(rollData)) return;
    const formula = rollData?.formula;
    if (!formula) return;

    if (rollEl.dataset?.tooltipText) {
      rollEl.dataset.tooltipText = formula;
    }

    const updateTooltip = () => {
      const tooltipFormula = rollEl.querySelector(".dice-tooltip .part-formula");
      const tooltipTotal = rollEl.querySelector(".dice-tooltip .part-total");
      const tooltipWrapper = rollEl.querySelector(".dice-tooltip .wrapper");
      let didUpdate = false;
      const shouldRewriteFormula = !tooltipMode || tooltipMode === "off";
      if (tooltipFormula && shouldRewriteFormula) {
        tooltipFormula.textContent = formula;
        didUpdate = true;
      }
      if (tooltipTotal || tooltipWrapper) {
        const finalTotal = Number(rollData?.total ?? tooltipTotal?.textContent ?? 0);
        const innerTotal = findInnerTotal(rollData);
        const baseValue = Number.isFinite(innerTotal)
          ? innerTotal
          : Number.isFinite(finalTotal)
            ? Math.floor(finalTotal / 1.5)
            : Number(tooltipTotal?.textContent ?? 0);
        if (Number.isFinite(baseValue)) {
          const totalValue = Number.isFinite(finalTotal)
            ? finalTotal
            : Math.floor(baseValue * 1.5);
          const systemHasNumericParts = hasSystemNumericParts(tooltipWrapper);
          if (!systemHasNumericParts) {
            const diceTotal = sumDicePartTotals(tooltipWrapper);
            const baseBonus = Number.isFinite(diceTotal) ? baseValue - diceTotal : null;
            const baseBonusLabel = inferBaseBonusLabel(formula);
            const baseRowUpdated = ensureNumericBonusRow(tooltipWrapper, baseBonus, baseBonusLabel, "nasBaseBonus");
            if (baseRowUpdated) didUpdate = true;
          }
          const empowerBonus = totalValue - baseValue;
          const rowUpdated = ensureNumericBonusRow(
            tooltipWrapper,
            empowerBonus,
            METAMAGIC_DEFINITION.name,
            "nasEmpowerBonus"
          );
          if (rowUpdated) didUpdate = true;
        }
      }
      return didUpdate;
    };

    updateTooltip();

    if (rollEl.dataset?.nasEmpowerObserver === "true") return;
    rollEl.dataset.nasEmpowerObserver = "true";

    const observer = new MutationObserver(() => {
      if (rollEl.__nasEmpowerObserverLock) return;
      rollEl.__nasEmpowerObserverLock = true;
      observer.disconnect();
      try {
        updateTooltip();
      } finally {
        rollEl.__nasEmpowerObserverLock = false;
        observer.observe(rollEl, { childList: true, subtree: true });
      }
    });

    observer.observe(rollEl, { childList: true, subtree: true });
  });
}
