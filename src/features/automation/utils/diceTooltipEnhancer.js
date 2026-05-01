import { MODULE } from "../../../common/module.js";

const TOOLTIP_WRAPPER_TARGET = "pf1.dice.RollPF.prototype.getTooltip";
const TOOLTIP_MODE_OFF = "off";
const TOOLTIP_MODE_LABELED = "labeled";
const TOOLTIP_MODE_DETAILED = "detailed";

let wrapperRegistered = false;

function getTooltipMode() {
  let value = TOOLTIP_MODE_OFF;
  try {
    value = game.settings.get(MODULE.ID, "enhancedDiceTooltipMode");
  } catch (_err) {
    return TOOLTIP_MODE_OFF;
  }
  if (value === TOOLTIP_MODE_LABELED || value === TOOLTIP_MODE_DETAILED) return value;
  return TOOLTIP_MODE_OFF;
}

function isMetamagicAutomationEnabled() {
  try {
    return Boolean(game.settings.get(MODULE.ID, "enableMetamagicAutomation"));
  } catch (_err) {
    return false;
  }
}

function isBlankFlavor(text) {
  const normalized = (text ?? "").toString().trim();
  if (!normalized) return true;
  return /^undefined$/i.test(normalized) || normalized === "-" || normalized === "—";
}

function localizeOrFallback(key, fallback) {
  const localized = game.i18n?.localize?.(key);
  if (!localized || localized === key) return fallback;
  return localized;
}

function hasCasterLevelPattern(formula) {
  const normalized = (formula ?? "").toString();
  if (!normalized) return false;
  return /@cl\b/i.test(normalized)
    || /min\s*\(\s*\d+\s*,\s*\d+\s*\)/i.test(normalized)
    || /floor\s*\(\s*\d+\s*\/\s*3\s*\)/i.test(normalized);
}

function getDetailedCasterLevelHint(formula) {
  const normalized = (formula ?? "").toString();
  if (!normalized) return "";
  const floorMatch = normalized.match(/floor\s*\(\s*\d+\s*\/\s*3\s*\)/i);
  if (floorMatch?.[0]) return floorMatch[0].replace(/\s+/g, " ");
  const minMatch = normalized.match(/min\s*\(\s*[^)]+\)/i);
  if (minMatch?.[0]) return minMatch[0].replace(/\s+/g, " ");
  if (/@cl\b/i.test(normalized)) return "@cl";
  return "";
}

function getCasterLevelBonusTotal(formula) {
  const normalized = (formula ?? "").toString();
  const floorMatch = normalized.match(/floor\s*\(\s*(\d+)\s*\/\s*3\s*\)/i);
  if (!floorMatch) return null;
  const casterLevel = Number(floorMatch[1]);
  if (!Number.isFinite(casterLevel)) return null;
  return Math.floor(casterLevel / 3);
}

function isInitiativeRoll(roll) {
  const flavor = (roll?.options?.flavor ?? "").toString().trim();
  return /^initiative check$/i.test(flavor);
}

function getInitiativeDexTiebreakerTotal(formula) {
  const normalized = (formula ?? "").toString();
  const match = normalized.match(/\(\(\s*([+-]?\d+(?:\.\d+)?)\s*\)\s*\/\s*100\s*\)\s*\[Tiebreaker\]/i);
  if (!match) return null;
  const value = Number(match[1]);
  return Number.isFinite(value) ? value : null;
}

function getPartTotalNumber(partEl) {
  const raw = (partEl.querySelector(".part-total")?.textContent ?? "").toString().trim();
  const value = Number(raw.replace(/^\+/, ""));
  return Number.isFinite(value) ? value : null;
}

function setPartFlavor(partEl, label, mode, hint = "") {
  const flavorEl = partEl.querySelector(".part-flavor");
  if (!flavorEl || !isBlankFlavor(flavorEl.textContent)) return;
  flavorEl.classList.remove("placeholder");
  if (flavorEl.dataset?.tooltip) {
    delete flavorEl.dataset.tooltip;
  }
  if (mode === TOOLTIP_MODE_DETAILED && hint) {
    flavorEl.textContent = `${label} (${hint})`;
    return;
  }
  flavorEl.textContent = label;
}

function appendSyntheticCasterLevelPart(template, { label, total, mode, hint = "" } = {}) {
  const wrapper = template.content.querySelector(".dice-tooltip .wrapper");
  if (!wrapper || !Number.isFinite(total)) return false;

  const section = document.createElement("section");
  section.classList.add("tooltip-part");
  const prefix = total >= 0 ? '<span class="part-prefix">+</span>' : "";
  const detail = mode === TOOLTIP_MODE_DETAILED && hint ? ` (${foundry.utils.escapeHTML(hint)})` : "";
  section.innerHTML = `
    <div class="dice">
      <header class="part-header flexrow">
        <span class="part-formula"></span>
        <span class="part-flavor">${foundry.utils.escapeHTML(label)}${detail}</span>
        <span class="part-total">${prefix}${total}</span>
      </header>
    </div>
  `;
  wrapper.append(section);
  return true;
}

export function enhanceDiceTooltipHtml(tooltipHtml, roll) {
  if (typeof tooltipHtml !== "string" || !tooltipHtml.trim()) return tooltipHtml;

  const mode = getTooltipMode();
  if (mode === TOOLTIP_MODE_OFF || !isMetamagicAutomationEnabled()) return tooltipHtml;

  const template = document.createElement("template");
  template.innerHTML = tooltipHtml;

  const partNodes = Array.from(template.content.querySelectorAll(".dice-tooltip .tooltip-part"));
  if (!partNodes.length) return tooltipHtml;

  const rollFormula = roll?.formula ?? "";
  const nasLabelContext = roll?.options?.nasLabelContext ?? {};
  const labelFormula = (nasLabelContext?.originalFormula ?? rollFormula ?? "").toString();
  const maximizeTransformed = Array.isArray(nasLabelContext?.transforms)
    && nasLabelContext.transforms.includes("maximizeSpell");
  const rollResultLabel = localizeOrFallback("PF1.RollResult", "Roll result");
  const casterLevelLabel = localizeOrFallback("PF1.CasterLevel", "Caster level");
  const bonusLabel = localizeOrFallback("PF1.Bonus", "Bonus");
  const dexterityLabel = localizeOrFallback("PF1.AbilityDex", "Dexterity");

  const hasCasterLevel = hasCasterLevelPattern(labelFormula);
  const initiativeDexTiebreakerTotal = isInitiativeRoll(roll)
    ? getInitiativeDexTiebreakerTotal(labelFormula)
    : null;
  let usedCasterLevelLabel = false;

  for (let idx = 0; idx < partNodes.length; idx += 1) {
    const partEl = partNodes[idx];
    const hasDiceResults = Boolean(partEl.querySelector(".dice-rolls li, .dice-rolls .roll"));
    const existingFlavor = (partEl.querySelector(".part-flavor")?.textContent ?? "").toString().trim();
    if (hasDiceResults) {
      setPartFlavor(partEl, rollResultLabel, mode);
      continue;
    }

    if (maximizeTransformed) {
      setPartFlavor(partEl, rollResultLabel, mode);
      continue;
    }

    if (
      Number.isFinite(initiativeDexTiebreakerTotal) &&
      getPartTotalNumber(partEl) === initiativeDexTiebreakerTotal
    ) {
      setPartFlavor(partEl, dexterityLabel, mode);
      continue;
    }

    if (hasCasterLevel && !usedCasterLevelLabel) {
      usedCasterLevelLabel = true;
      const hint = mode === TOOLTIP_MODE_DETAILED ? getDetailedCasterLevelHint(labelFormula) : "";
      setPartFlavor(partEl, casterLevelLabel, mode, hint);
      continue;
    }

    setPartFlavor(partEl, bonusLabel, mode);
  }

  if (hasCasterLevel && !usedCasterLevelLabel) {
    const hint = mode === TOOLTIP_MODE_DETAILED ? getDetailedCasterLevelHint(labelFormula) : "";
    appendSyntheticCasterLevelPart(template, {
      label: casterLevelLabel,
      total: getCasterLevelBonusTotal(labelFormula),
      mode,
      hint
    });
  }

  return template.innerHTML;
}

export function registerDiceTooltipEnhancer() {
  if (wrapperRegistered) return;
  if (!globalThis.libWrapper?.register) return;

  globalThis.libWrapper.register(
    MODULE.ID,
    TOOLTIP_WRAPPER_TARGET,
    async function wrappedGetTooltip(wrapped, ...args) {
      const tooltipHtml = await wrapped(...args);
      return enhanceDiceTooltipHtml(tooltipHtml, this);
    },
    "WRAPPER"
  );

  wrapperRegistered = true;
}
