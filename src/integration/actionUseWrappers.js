import { MODULE } from '../common/module.js';
import { handleBuffAutomation } from '../features/automation/buffs/buffs.js';
import { collectSpellActionData } from '../features/automation/utils/spellActionData.js';
import { commitPersistentFeatureStatesFromOptions } from '../features/automation/utils/attackDialogControls.js';
import { applyMetamagicSelections } from '../features/automation/metamagic/applyMetamagic.js';
import { METAMAGIC_DEFINITIONS } from '../features/automation/metamagic/metamagic.js';
import { applyActionUseOverrides } from '../features/automation/utils/actionUseOverrides.js';
import {
  applyMaleficiumPostMetamagic,
  applyEldritchResearcherPostMetamagic,
  prepareMaleficiumContext,
  prepareEldritchResearcherContext,
  prepareSpellPerfectionContext,
  prepareSpontaneousMetafocusContext,
  prepareExtendedScryingContext,
  prepareMaskFocusContext
} from '../features/automation/metamagic/feats/index.js';
import {
  prepareMagicalLineageContext,
  prepareWayangSpellhunterContext
} from '../features/automation/metamagic/traits/index.js';
import {
  createGrappleCmbAttackEntry,
  isGrappleCmbAttack,
  isGrappleSelected,
} from '../features/automation/conditions/grappled/grappled.js';
import {
  handleMirrorImageCast,
  resolveMirrorImagesForActionUse
} from '../features/automation/buffs/mirrorImage.js';
import {
  contextHasSavingThrow,
  resolveFeatSaveDcBase,
} from '../features/automation/utils/saveDcUtils.js';

function shouldHandleMetamagic(actionUse) {
  const itemType = actionUse?.item?.type;
  const itemSubType = actionUse?.item?.subType;
  if (!game.settings.get(MODULE.ID, "enableMetamagicAutomation")) return false;
  return itemType === "spell" || itemType === "consumable" || (itemType === "feat" && itemSubType === "classFeat");
}

function shouldHandleAutomaticBuffs(actionUse) {
  const itemType = actionUse?.item?.type;
  const itemSubType = actionUse?.item?.subType;
  const useCustomLogic = (
    itemType === "spell" ||
    itemType === "consumable" ||
    (itemType === "feat" && itemSubType === "classFeat")
  );
  return useCustomLogic && game.settings.get(MODULE.ID, "automaticBuffs");
}

function buildRangeOverride(actionUse) {
  const rangeUnits = actionUse.action?.range?.units;
  const increments = actionUse.action?.range?.maxIncrements;
  const rollData = actionUse.shared.rollData;
  const baseRange = actionUse.action?.getRange?.({ type: "single", rollData });
  const minRange = actionUse.action?.getRange?.({ type: "min", rollData });
  const maxRange = actionUse.action?.getRange?.({ type: "max", rollData });
  return {
    base: baseRange != null ? pf1.utils.convertDistanceBack(baseRange)[0] : null,
    min: minRange != null ? pf1.utils.convertDistanceBack(minRange)[0] : null,
    max: maxRange != null ? pf1.utils.convertDistanceBack(maxRange)[0] : null,
    units: rangeUnits ?? "",
    increments: increments ?? 1,
  };
}

function escapeRegExp(value) {
  return (value ?? "").toString().replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function sanitizeSpellDescriptionComponents(description, { removeVerbal = false, removeSomatic = false } = {}) {
  if (!description || typeof description !== "string") return description;
  if (!removeVerbal && !removeSomatic) return description;

  const componentsLabel = game.i18n.localize("PF1.Components");
  const verbalLabel = pf1?.config?.spellComponents?.verbal ?? "V";
  const somaticLabel = game.i18n.localize(pf1?.config?.spellComponents?.somatic ?? "PF1.SpellComponents.Type.somatic.Abbr");
  const componentsRegex = new RegExp(
    `(<strong[^>]*>\\s*${escapeRegExp(componentsLabel)}\\s*<\\/strong>\\s*&nbsp;)([^<]*)(<br\\s*\\/?>)`,
    "i"
  );
  const match = description.match(componentsRegex);
  if (!match) return description;

  const before = match[2];
  const nextComponents = before
    .split(",")
    .map((part) => part.trim())
    .filter((part) => {
      if (part.length === 0) return false;
      if (removeVerbal && part === verbalLabel) return false;
      if (removeSomatic && part === somaticLabel) return false;
      return true;
    });
  if (nextComponents.length === before.split(",").map((part) => part.trim()).filter(Boolean).length) {
    return description;
  }
  if (!nextComponents.length) {
    return description.replace(match[0], "");
  }
  return description.replace(match[0], `${match[1]}${nextComponents.join(", ")}${match[3]}`);
}

function coerceNumericSaveDc(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value === "string" && value.trim().length > 0) {
    const parsed = Number(value.trim());
    if (Number.isFinite(parsed)) return parsed;
  }
  return null;
}

const METAMAGIC_FIXED_SLOT_COST = new Map([
  ["Dazing Spell", 3],
  ["Empower Spell", 2],
  ["Enlarge Spell", 1],
  ["Extend Spell", 1],
  ["Intensified Spell", 1],
  ["Maximize Spell", 3],
  ["Persistent Spell", 2],
  ["Quicken Spell", 4],
  ["Selective Spell", 1],
  ["Silent Spell", 1],
  ["Still Spell", 1],
]);

const METAMAGIC_PREFIX_BY_NAME = new Map(
  METAMAGIC_DEFINITIONS.map((definition) => [definition.name, definition.prefix]).filter(([, prefix]) => Boolean(prefix))
);

function getAppliedMetamagicNames(context) {
  if (Array.isArray(context?.metamagic?.applied) && context.metamagic.applied.length > 0) {
    return context.metamagic.applied;
  }
  if (Array.isArray(context?.metamagicNames) && context.metamagicNames.length > 0) {
    return context.metamagicNames;
  }
  return [];
}

function getBaseSpellLevel(actionUse, context) {
  const fromContext = Number(context?.spellLevel?.original);
  if (Number.isFinite(fromContext)) return fromContext;
  const fromItem = Number(actionUse?.item?.system?.level);
  if (Number.isFinite(fromItem)) return fromItem;
  return 0;
}

function getMetamagicSlotCost(name, actionUse, context) {
  if (name === "Extend Spell" && context?.metamagic?.extendSlotWaivedByMaskFocus === true) {
    return 0;
  }
  if (name === "Reach Spell") {
    const steps = Number(context?.metamagicOptions?.reachSpellSteps ?? 1);
    return Number.isFinite(steps) && steps > 0 ? steps : 1;
  }
  if (name === "Heighten Spell") {
    const baseSpellLevel = getBaseSpellLevel(actionUse, context);
    const heightenLevel = Number(context?.metamagic?.heightenLevel ?? context?.metamagicOptions?.heightenSpellLevel ?? 0);
    if (!Number.isFinite(heightenLevel)) return 0;
    return Math.max(0, heightenLevel - baseSpellLevel);
  }
  return Number(METAMAGIC_FIXED_SLOT_COST.get(name) ?? 0);
}

function maybeApplyMetamagicChatCardNamePrefix(actionUse, context) {
  const mode = game.settings.get(MODULE.ID, "metamagicChatCardNameMode");
  if (mode !== "highest") return;

  const currentName = actionUse?.shared?.templateData?.name;
  if (typeof currentName !== "string" || currentName.trim().length === 0) return;

  context.metamagic ??= {};
  const cachedName = context.metamagic.chatCardPrefixSourceName;
  if (typeof cachedName === "string" && cachedName.length > 0) {
    const cachedPrefix = METAMAGIC_PREFIX_BY_NAME.get(cachedName);
    if (cachedPrefix) {
      const prefixed = `${cachedPrefix} `;
      if (!currentName.startsWith(prefixed)) {
        actionUse.shared.templateData.name = `${cachedPrefix} ${currentName}`;
      }
    }
    return;
  }

  const appliedNames = getAppliedMetamagicNames(context);
  if (!appliedNames.length) return;

  let maxCost = Number.NEGATIVE_INFINITY;
  let candidates = [];
  for (const name of appliedNames) {
    const cost = getMetamagicSlotCost(name, actionUse, context);
    if (cost > maxCost) {
      maxCost = cost;
      candidates = [name];
    } else if (cost === maxCost) {
      candidates.push(name);
    }
  }

  if (!candidates.length) return;
  const selectedName = candidates[Math.floor(Math.random() * candidates.length)];
  context.metamagic.chatCardPrefixSourceName = selectedName;
  const prefix = METAMAGIC_PREFIX_BY_NAME.get(selectedName);
  if (!prefix) return;
  if (!currentName.startsWith(`${prefix} `)) {
    actionUse.shared.templateData.name = `${prefix} ${currentName}`;
  }
}

function findOrCreateInfoPropertyGroup(properties) {
  const localizedInfoHeader = game.i18n.localize("PF1.InfoShort");
  let group = properties.find((entry) => (
    entry
    && typeof entry === "object"
    && Array.isArray(entry.value)
    && (entry.header === localizedInfoHeader || entry.css === "common-notes")
  ));
  if (!group) {
    group = {
      header: localizedInfoHeader,
      css: "common-notes",
      value: []
    };
    properties.push(group);
  }
  group.value = Array.isArray(group.value) ? group.value : [];
  return group;
}

function addInfoTag(infoGroup, label) {
  if (!infoGroup || typeof label !== "string") return;
  const text = label.trim();
  if (!text) return;
  if (infoGroup.value.some((entry) => entry?.text === text)) return;
  infoGroup.value.push({ text });
}

function addInfoNames(infoGroup, names) {
  if (!Array.isArray(names)) return;
  for (const name of names) {
    if (typeof name !== "string") continue;
    addInfoTag(infoGroup, name);
  }
}

function buildActivationLabel(activation) {
  if (!activation || typeof activation !== "object") return "";
  const activationType = (activation.type ?? "nonaction").toString();
  if (!activationType) return "";

  const isUnchainedActionEconomy = game.settings.get("pf1", "unchainedActionEconomy");
  const activationTypes = isUnchainedActionEconomy
    ? pf1.config.abilityActivationTypes_unchained
    : pf1.config.abilityActivationTypes;
  const activationTypesPlural = isUnchainedActionEconomy
    ? pf1.config.abilityActivationTypesPlurals_unchained
    : pf1.config.abilityActivationTypesPlurals;

  if (activationType === "special") {
    return activationTypes.special ?? "";
  }

  const rawCost = Number(activation.cost);
  const hasNumericCost = Number.isFinite(rawCost) && rawCost > 0;
  if (hasNumericCost && rawCost > 1 && activationTypesPlural?.[activationType]) {
    return [String(rawCost), activationTypesPlural[activationType]].filterJoin(" ");
  }

  const prefixCost = ["minute", "hour", "action"].includes(activationType) && hasNumericCost
    ? String(rawCost)
    : "";
  return [prefixCost, activationTypes?.[activationType]].filterJoin(" ");
}

function updateChatInfoActivationLabel(templateData, beforeActivation, afterActivation) {
  if (!templateData || !Array.isArray(templateData.properties)) return { changed: false, beforeLabel: "", afterLabel: "" };
  const beforeLabel = buildActivationLabel(beforeActivation).trim();
  const afterLabel = buildActivationLabel(afterActivation).trim();
  if (!beforeLabel || !afterLabel || beforeLabel === afterLabel) {
    return { changed: false, beforeLabel, afterLabel };
  }

  const infoGroup = findOrCreateInfoPropertyGroup(templateData.properties);
  if (!Array.isArray(infoGroup?.value) || infoGroup.value.length === 0) {
    return { changed: false, beforeLabel, afterLabel };
  }

  let replaced = false;
  for (const entry of infoGroup.value) {
    if (typeof entry === "string") {
      if (entry.trim() === beforeLabel) {
        const index = infoGroup.value.indexOf(entry);
        if (index >= 0) infoGroup.value[index] = afterLabel;
        replaced = true;
        break;
      }
      continue;
    }
    if (entry && typeof entry === "object" && typeof entry.text === "string" && entry.text.trim() === beforeLabel) {
      entry.text = afterLabel;
      replaced = true;
      break;
    }
  }

  return { changed: replaced, beforeLabel, afterLabel };
}

const EXTEND_SPELL_METAMAGIC_NAME = "Extend Spell";
const MASK_FOCUS_PAREN_NOTE = "(mask focus)";

/**
 * Append ", mask focus" inside existing focus/M/DF/DF parentheses on the Components line, or add
 * ", F (mask focus)" when none apply. Matches PF1 divineFocus: 1=DF, 2=MDF, 3=FDF (see item-spell.mjs).
 */
function patchMaskFocusComponentsInDescription(html, item, context) {
  if (typeof html !== "string" || !html || !item) return { changed: false, html };
  if (context?.featEffects?.maskFocus?.active !== true) return { changed: false, html };
  const applied = context?.metamagic?.applied ?? [];
  if (!Array.isArray(applied) || !applied.includes(EXTEND_SPELL_METAMAGIC_NAME)) {
    return { changed: false, html };
  }

  const componentsLabel = game.i18n.localize("PF1.Components");
  const componentsRegex = new RegExp(
    `(<strong[^>]*>\\s*${escapeRegExp(componentsLabel)}\\s*<\\/strong>\\s*(?:&nbsp;|\\u00a0|\\s)*)([^<]*)(<br\\s*\\/?>)`,
    "i"
  );
  const match = html.match(componentsRegex);
  if (!match) return { changed: false, html };

  let body = match[2];
  if (body.toLowerCase().includes(MASK_FOCUS_PAREN_NOTE)) return { changed: false, html };

  const focusAbbr = game.i18n.localize(pf1?.config?.spellComponents?.focus ?? "PF1.SpellComponents.Type.focus.Abbr");
  const dfAbbr = game.i18n.localize(pf1?.config?.spellComponents?.divineFocus ?? "PF1.SpellComponents.Type.divineFocus.Abbr");
  const matAbbr = game.i18n.localize(pf1?.config?.spellComponents?.material ?? "PF1.SpellComponents.Type.material.Abbr");
  const fe = escapeRegExp(focusAbbr);
  const dfe = escapeRegExp(dfAbbr);
  const me = escapeRegExp(matAbbr);

  const dfRaw = Number(item?.system?.components?.divineFocus ?? 0);
  const df = Number.isFinite(dfRaw) ? dfRaw : 0;

  const appendInsideOutermostParens = (segment) => {
    const open = segment.indexOf("(");
    const close = segment.lastIndexOf(")");
    if (open === -1 || close < open) return segment;
    const inner = segment.slice(open + 1, close);
    if (inner.toLowerCase().includes("mask focus")) return segment;
    return `${segment.slice(0, close)}, mask focus${segment.slice(close)}`;
  };

  const replaceLastGlobal = (source, patternSource) => {
    const re = new RegExp(patternSource, "gi");
    let last = null;
    let m;
    while ((m = re.exec(source)) !== null) {
      last = { full: m[0], index: m.index };
    }
    if (!last) return null;
    const updated = appendInsideOutermostParens(last.full);
    if (updated === last.full) return null;
    return source.slice(0, last.index) + updated + source.slice(last.index + last.full.length);
  };

  let nextBody = replaceLastGlobal(body, `${fe}\\s*/\\s*${dfe}\\s*\\([^)]*\\)`);
  if (!nextBody) nextBody = replaceLastGlobal(body, `${fe}\\s*\\([^)]*\\)`);
  if (!nextBody && df === 2) {
    nextBody = replaceLastGlobal(body, `${me}\\s*/\\s*${dfe}\\s*\\([^)]*\\)`);
  }
  if (!nextBody && df === 1) {
    const dfOnly = new RegExp(`(?<![\\/])${dfe}\\s*$`, "i");
    const trimmed = body.trimEnd();
    if (dfOnly.test(trimmed)) {
      nextBody = trimmed.replace(dfOnly, `${dfAbbr} ${MASK_FOCUS_PAREN_NOTE}`);
    }
  }
  if (!nextBody) {
    const trimmed = body.trimEnd();
    const sep = trimmed.length ? ", " : "";
    nextBody = `${trimmed}${sep}${focusAbbr} ${MASK_FOCUS_PAREN_NOTE}`;
  }

  if (nextBody === body) return { changed: false, html };
  return {
    changed: true,
    html: html.replace(componentsRegex, `$1${nextBody}$3`)
  };
}

function patchMaskFocusDurationInDescription(html, context) {
  if (typeof html !== "string" || !html) return { changed: false, html };
  const mf = context?.duration?.maskFocusSelf;
  if (!mf || context?.featEffects?.maskFocus?.active !== true) return { changed: false, html };

  const durationLabel = game.i18n.localize("PF1.Duration");
  const unitLabel = pf1?.config?.timePeriods?.[mf.units] ?? mf.units ?? "";
  const suffix = game.i18n.format("NAS.metamagic.maskFocus.durationOnSelfSuffix", {
    value: String(mf.extendedSelfTotal),
    units: unitLabel
  });

  const escapedLabel = escapeRegExp(durationLabel);
  const pattern = new RegExp(
    `(<strong[^>]*>\\s*${escapedLabel}\\s*<\\/strong>\\s*(?:&nbsp;|\\u00a0|\\s)*)(<span[^>]*>)([\\s\\S]*?)(<\\/span>)(\\s*<br\\s*\\/?>)`,
    "i"
  );
  const match = html.match(pattern);
  if (!match) return { changed: false, html };

  const innerRaw = match[3];
  if (innerRaw.includes(suffix.trim())) return { changed: false, html };

  const newInner = `${innerRaw.trimEnd()} ${suffix}`;
  return {
    changed: true,
    html: html.replace(pattern, `$1$2${newInner}$4$5`)
  };
}

function updateDescriptionActivationLabel(html, beforeLabel, afterLabel) {
  if (typeof html !== "string" || !html) return { changed: false, html };
  const before = (beforeLabel ?? "").trim();
  const after = (afterLabel ?? "").trim();
  if (!before || !after || before === after) return { changed: false, html };

  const castingTimeLabel = game.i18n.localize("PF1.CastingTime");
  const pattern = new RegExp(
    `(<strong[^>]*>\\s*${escapeRegExp(castingTimeLabel)}\\s*<\\/strong>\\s*(?:&nbsp;|\\u00a0|\\s)*)${escapeRegExp(before)}(\\s*<br\\s*\\/?>)`,
    "i"
  );
  if (!pattern.test(html)) return { changed: false, html };
  return {
    changed: true,
    html: html.replace(pattern, `$1${after}$2`)
  };
}

function buildChatDurationDisplay(context) {
  const duration = context?.duration;
  if (!duration) return null;
  const units = (duration.units ?? "").toString();
  let display = null;
  if (!units) return null;
  if (units === "spec") {
    display = duration.value ?? "";
  } else if (["seeText", "inst", "perm"].includes(units)) {
    display = pf1?.config?.timePeriods?.[units] ?? units;
  } else if (["turn", "round", "minute", "hour", "day", "month", "year"].includes(units)) {
    const unit = pf1?.config?.timePeriods?.[units] ?? units;
    const value = units === "turn"
      ? 1
      : (context?.duration?.evaluated?.total ?? context?.duration?.value ?? "");
    display = game.i18n.format("PF1.Time.Format", { value, unit });
  }
  if (!display) return null;
  if (units !== "spec") {
    if (duration.dismiss) display += ` ${game.i18n.localize("PF1.DismissableMark")}`;
    if (duration.concentration) {
      display = game.i18n.format("PF1.ConcentationDuration", { duration: display });
    }
  }
  return display;
}

function patchDurationInDescription(html, durationDisplay) {
  if (typeof html !== "string" || !html || !durationDisplay) return { changed: false, html };
  const durationLabel = game.i18n.localize("PF1.Duration");
  const pattern = new RegExp(
    `(<strong[^>]*>\\s*${escapeRegExp(durationLabel)}\\s*<\\/strong>\\s*(?:&nbsp;|\\u00a0|\\s)*)(<span[^>]*>)([\\s\\S]*?)(<\\/span>)(\\s*<br\\s*\\/?>)`,
    "i"
  );
  if (!pattern.test(html)) return { changed: false, html };
  return {
    changed: true,
    html: html.replace(pattern, `$1$2${durationDisplay}$4$5`)
  };
}

function patchDurationInInfoProperties(properties, durationDisplay) {
  if (!Array.isArray(properties) || !durationDisplay) return false;
  const localizedInfoHeader = game.i18n.localize("PF1.InfoShort");
  const infoGroups = properties.filter((entry) => (
    entry
    && typeof entry === "object"
    && Array.isArray(entry.value)
    && (entry.header === "Info" || entry.header === localizedInfoHeader || entry.css === "common-notes")
  ));
  if (!infoGroups.length) return false;
  const unitLabels = ["round", "minute", "hour", "day", "month", "year", "turn", "inst", "perm", "seeText"]
    .map((key) => pf1?.config?.timePeriods?.[key])
    .filter((value) => typeof value === "string" && value.trim().length > 0)
    .map((value) => value.toLowerCase());
  const matches = [];
  for (const group of infoGroups) {
    group.value.forEach((entry, index) => {
      const text = (entry?.text ?? "").toString().trim();
      if (!text) return;
      if (text.startsWith("Charges:") || text.startsWith("Charge Cost:")) return;
      const lower = text.toLowerCase();
      if (!unitLabels.some((label) => lower.includes(label))) return;
      matches.push({ group, index });
    });
  }
  if (!matches.length) return false;
  const primary = matches[0];
  primary.group.value[primary.index].text = durationDisplay;
  for (let i = matches.length - 1; i >= 1; i -= 1) {
    const { group, index } = matches[i];
    group.value.splice(index, 1);
  }
  return true;
}

function applyMetamagicChatCardFooters(actionUse, context) {
  const templateData = actionUse?.shared?.templateData;
  if (!templateData) return;
  if (!Array.isArray(templateData.properties)) {
    templateData.properties = [];
  }
  const properties = templateData.properties;
  const infoGroup = findOrCreateInfoPropertyGroup(properties);

  const hasSave = contextHasSavingThrow(context, actionUse);
  const baseSnap = coerceNumericSaveDc(context?.save?.baseDc);
  const coercedFinal = coerceNumericSaveDc(context?.save?.dc ?? actionUse?.shared?.saveDC);
  let displayFinal = NaN;
  if (hasSave) {
    const resolvedFinal = Number.isFinite(coercedFinal)
      ? coercedFinal
      : resolveFeatSaveDcBase(actionUse, context);
    displayFinal = Number.isFinite(coercedFinal)
      ? coercedFinal
      : Number.isFinite(resolvedFinal)
        ? resolvedFinal
        : NaN;
  }

  const dcChangedFromSnapshot =
    hasSave &&
    Number.isFinite(baseSnap) &&
    Number.isFinite(displayFinal) &&
    baseSnap !== displayFinal;
  const explicitNumericDcWithoutSnapshot =
    hasSave &&
    Number.isFinite(coercedFinal) &&
    !Number.isFinite(baseSnap);
  const shouldReplaceDcFootnotes = dcChangedFromSnapshot || explicitNumericDcWithoutSnapshot;

  if (shouldReplaceDcFootnotes) {
    infoGroup.value = infoGroup.value.filter((entry) => {
      const text = (entry?.text ?? "").toString().trim();
      if (!text) return true;
      if (/^Base\s+DC\s+\d+/i.test(text)) return false;
      if (/^DC\s+\d+/i.test(text)) return false;
      return true;
    });
  }

  if (dcChangedFromSnapshot) {
    addInfoTag(infoGroup, `Base DC ${baseSnap}`);
    addInfoTag(infoGroup, `DC ${displayFinal}`);
  } else if (explicitNumericDcWithoutSnapshot) {
    addInfoTag(infoGroup, `DC ${coercedFinal}`);
  }

  const appliedMetamagics = getAppliedMetamagicNames(context);
  addInfoNames(infoGroup, appliedMetamagics);
  addInfoNames(infoGroup, context?.metamagic?.activeFeatureLabels);
}

async function consumePendingMaskFocusUse(actionUse) {
  const pending = actionUse?.shared?.nasPendingMaskFocusUse;
  if (!pending) return;
  delete actionUse.shared.nasPendingMaskFocusUse;

  if (actionUse?.shared?.reject === true || actionUse?.shared?.scriptData?.reject === true) return;

  const actor = actionUse?.actor ?? actionUse?.token?.actor ?? null;
  if (!actor) return;

  let featureItem = null;
  const pendingUuid = (pending?.itemUuid ?? "").toString();
  if (pendingUuid && typeof fromUuid === "function") {
    try {
      featureItem = await fromUuid(pendingUuid);
    } catch (_error) {
      featureItem = null;
    }
  }

  if (!featureItem) return;
  if (featureItem.actor?.id && featureItem.actor.id !== actor.id) return;

  const uses = featureItem?.system?.uses;
  if (!uses || !uses.per) {
    ui.notifications.warn(`${featureItem.name}: missing uses metadata, skipping Mask Focus consumption.`);
    return;
  }

  const remaining = Number(uses.value ?? 0);
  if (!Number.isFinite(remaining) || remaining <= 0) {
    return;
  }

  const tokenDocument = actionUse?.token?.document ?? null;
  try {
    await featureItem.use?.({ skipDialog: true, chatMessage: false, token: tokenDocument });
  } catch (_error) {
    const fallbackRemaining = Number(featureItem?.system?.uses?.value ?? 0);
    if (Number.isFinite(fallbackRemaining) && fallbackRemaining > 0) {
      await featureItem.update({ "system.uses.value": fallbackRemaining - 1 });
    }
  }
}

const ARCANE_RESERVOIR_ARCANIST_SOURCE = "Compendium.pf1.class-abilities.Item.CtDtLshBC8pc64JV";
const ARCANE_RESERVOIR_EXPLOITER_SOURCE =
  "Compendium.pf1e-archetypes.pf-arch-features.Item.tWiV2EJuWxUHBFCd";
const PHRENIC_POOL_SOURCE = "Compendium.pf1.class-abilities.Item.tFy3rxyljSq56HSg";

async function consumePendingPhrenicPoolSpend(actionUse) {
  const pending = actionUse?.shared?.nasPendingPhrenicPoolSpend;
  if (!pending) return;
  delete actionUse.shared.nasPendingPhrenicPoolSpend;

  if (actionUse?.shared?.reject === true || actionUse?.shared?.scriptData?.reject === true) return;

  const actor = actionUse?.actor ?? actionUse?.token?.actor ?? null;
  if (!actor) return;

  let poolItem = null;
  const pendingUuid = (pending?.itemUuid ?? "").toString();
  if (pendingUuid && typeof fromUuid === "function") {
    try {
      poolItem = await fromUuid(pendingUuid);
    } catch (_error) {
      poolItem = null;
    }
  }

  if (!poolItem) {
    poolItem = actor.items.find((item) => {
      const source = item?._stats?.compendiumSource ?? "";
      return source === PHRENIC_POOL_SOURCE;
    }) ?? null;
  }

  if (!poolItem) return;
  if (poolItem.actor?.id && poolItem.actor.id !== actor.id) return;

  const uses = poolItem?.system?.uses;
  if (!uses || !uses.per) {
    ui.notifications.warn(`${poolItem.name}: missing uses metadata, skipping phrenic pool spend.`);
    return;
  }

  const debitRaw = Number(pending?.debitCount ?? 1);
  const debitCount = Number.isFinite(debitRaw) && debitRaw > 0 ? Math.floor(debitRaw) : 1;
  const remaining = Number(uses.value ?? 0);
  if (!Number.isFinite(remaining) || remaining < debitCount) return;

  const tokenDocument = actionUse?.token?.document ?? null;
  for (let i = 0; i < debitCount; i += 1) {
    try {
      await poolItem.use?.({ skipDialog: true, chatMessage: false, token: tokenDocument });
    } catch (_error) {
      const fallbackRemaining = Number(poolItem?.system?.uses?.value ?? 0);
      if (Number.isFinite(fallbackRemaining) && fallbackRemaining > 0) {
        await poolItem.update({ "system.uses.value": fallbackRemaining - 1 });
      }
    }
  }
}

async function consumePendingArcaneReservoirSpend(actionUse) {
  const pending = actionUse?.shared?.nasPendingArcaneReservoirSpend;
  if (!pending) return;
  delete actionUse.shared.nasPendingArcaneReservoirSpend;

  if (actionUse?.shared?.reject === true || actionUse?.shared?.scriptData?.reject === true) return;

  const actor = actionUse?.actor ?? actionUse?.token?.actor ?? null;
  if (!actor) return;

  let reservoirItem = null;
  const pendingUuid = (pending?.itemUuid ?? "").toString();
  if (pendingUuid && typeof fromUuid === "function") {
    try {
      reservoirItem = await fromUuid(pendingUuid);
    } catch (_error) {
      reservoirItem = null;
    }
  }

  if (!reservoirItem) {
    reservoirItem = actor.items.find((item) => {
      const source = item?._stats?.compendiumSource ?? "";
      return source === ARCANE_RESERVOIR_ARCANIST_SOURCE || source === ARCANE_RESERVOIR_EXPLOITER_SOURCE;
    }) ?? null;
  }

  if (!reservoirItem) return;
  if (reservoirItem.actor?.id && reservoirItem.actor.id !== actor.id) return;

  const uses = reservoirItem?.system?.uses;
  if (!uses || !uses.per) {
    ui.notifications.warn(`${reservoirItem.name}: missing uses metadata, skipping arcane reservoir spend.`);
    return;
  }

  const debitRaw = Number(pending?.debitCount ?? 1);
  const debitCount = Number.isFinite(debitRaw) && debitRaw > 0 ? Math.floor(debitRaw) : 1;

  const remaining = Number(uses.value ?? 0);
  if (!Number.isFinite(remaining) || remaining < debitCount) {
    return;
  }

  const tokenDocument = actionUse?.token?.document ?? null;
  for (let i = 0; i < debitCount; i += 1) {
    try {
      await reservoirItem.use?.({ skipDialog: true, chatMessage: false, token: tokenDocument });
    } catch (_error) {
      const fallbackRemaining = Number(reservoirItem?.system?.uses?.value ?? 0);
      if (Number.isFinite(fallbackRemaining) && fallbackRemaining > 0) {
        await reservoirItem.update({ "system.uses.value": fallbackRemaining - 1 });
      }
    }
  }
}

async function consumePendingMetamagicFeatureUse(actionUse) {
  const pending = actionUse?.shared?.nasPendingMetamagicUse;
  if (!pending) return;
  delete actionUse.shared.nasPendingMetamagicUse;

  if (actionUse?.shared?.reject === true || actionUse?.shared?.scriptData?.reject === true) return;

  const actor = actionUse?.actor ?? actionUse?.token?.actor ?? null;
  if (!actor) return;

  let featureItem = null;
  const pendingUuid = (pending?.itemUuid ?? "").toString();
  if (pendingUuid && typeof fromUuid === "function") {
    try {
      featureItem = await fromUuid(pendingUuid);
    } catch (_error) {
      featureItem = null;
    }
  }

  if (!featureItem && pending?.featureId) {
    if (pending.featureId === "metamagicAdept") {
      featureItem = actor.items.find((item) => {
        const source = item?._stats?.compendiumSource ?? "";
        return source === "Compendium.pf1.class-abilities.Item.EwV7Db8W6ww3rtd0";
      }) ?? null;
    }
    if (pending.featureId === "metamagicMastery") {
      featureItem = actor.items.find((item) => {
        const source = item?._stats?.compendiumSource ?? "";
        return source === "Compendium.pf1.class-abilities.Item.L4sgpJ5DxgYddhem";
      }) ?? null;
    }
  }

  if (!featureItem) return;
  if (featureItem.actor?.id && featureItem.actor.id !== actor.id) return;

  const uses = featureItem?.system?.uses;
  if (!uses || !uses.per) {
    ui.notifications.warn(`${featureItem.name}: missing uses metadata, skipping consumption.`);
    return;
  }

  const debitRaw = Number(pending?.debitCount ?? 1);
  const debitCount = Number.isFinite(debitRaw) && debitRaw > 0 ? Math.floor(debitRaw) : 1;

  const remaining = Number(uses.value ?? 0);
  if (!Number.isFinite(remaining) || remaining < debitCount) {
    return;
  }

  const tokenDocument = actionUse?.token?.document ?? null;
  for (let i = 0; i < debitCount; i += 1) {
    try {
      await featureItem.use?.({ skipDialog: true, chatMessage: false, token: tokenDocument });
    } catch (_error) {
      const fallbackRemaining = Number(featureItem?.system?.uses?.value ?? 0);
      if (Number.isFinite(fallbackRemaining) && fallbackRemaining > 0) {
        await featureItem.update({ "system.uses.value": fallbackRemaining - 1 });
      }
    }
  }
}

async function resolvePendingFeatureItem(actor, itemUuid) {
  const pendingUuid = (itemUuid ?? "").toString();
  if (!pendingUuid || typeof fromUuid !== "function") return null;
  try {
    const item = await fromUuid(pendingUuid);
    if (!item) return null;
    if (item.actor?.id && item.actor.id !== actor?.id) return null;
    return item;
  } catch (_error) {
    return null;
  }
}

async function consumePendingTraitFeatureUses(actionUse) {
  const pending = actionUse?.shared?.nasPendingTraitUses;
  if (!Array.isArray(pending) || !pending.length) return;
  delete actionUse.shared.nasPendingTraitUses;

  if (actionUse?.shared?.reject === true || actionUse?.shared?.scriptData?.reject === true) return;

  const actor = actionUse?.actor ?? actionUse?.token?.actor ?? null;
  if (!actor) return;
  const tokenDocument = actionUse?.token?.document ?? null;

  for (const entry of pending) {
    const traitItem = await resolvePendingFeatureItem(actor, entry?.itemUuid);
    if (!traitItem) continue;

    if ((entry?.mode ?? "use") === "displayCard") {
      try {
        await traitItem.displayCard?.({}, { token: tokenDocument });
      } catch (_error) {
        // Best-effort only.
      }
      continue;
    }

    const uses = traitItem?.system?.uses;
    if (!uses || !uses.per) {
      try {
        await traitItem.displayCard?.({}, { token: tokenDocument });
      } catch (_error) {
      }
      continue;
    }

    const remaining = Number(uses.value ?? 0);
    if (!Number.isFinite(remaining) || remaining <= 0) continue;

    try {
      await traitItem.use?.({ skipDialog: true, chatMessage: true, token: tokenDocument });
    } catch (_error) {
      const fallbackRemaining = Number(traitItem?.system?.uses?.value ?? 0);
      if (Number.isFinite(fallbackRemaining) && fallbackRemaining > 0) {
        await traitItem.update({ "system.uses.value": fallbackRemaining - 1 });
      }
    }
  }
}

export function registerActionUseWrapper() {
  if (!game.modules.get("lib-wrapper")?.active) {
    ui.notifications.error(`${MODULE.NAME} requires the 'libWrapper' module. Please install and activate it.`);
    return;
  }

  libWrapper.register(
    MODULE.ID,
    "pf1.actionUse.ActionUse.prototype.process",
    async function (wrapped, ...args) {
      try {
        const result = await wrapped.apply(this, args);
        await consumePendingMetamagicFeatureUse(this);
        await consumePendingPhrenicPoolSpend(this);
        await consumePendingArcaneReservoirSpend(this);
        await consumePendingMaskFocusUse(this);
        await consumePendingTraitFeatureUses(this);
        return result;
      } finally {
        const restore = this?.shared?.nasRestoreOverrides;
        if (typeof restore === "function") {
          restore();
          delete this.shared.nasRestoreOverrides;
        }
      }
    },
    "MIXED"
  );

  libWrapper.register(
    MODULE.ID,
    "pf1.actionUse.ActionUse.prototype.alterRollData",
    async function (wrapped, ...args) {
      const result = await wrapped.apply(this, args);
      if (!shouldHandleMetamagic(this)) return result;
      try {
        this.shared.nasSpellContext = await collectSpellActionData(this);
        commitPersistentFeatureStatesFromOptions(this?.actor ?? this?.token?.actor ?? null, this.shared.nasSpellContext?.metamagicOptions ?? {});
        await prepareExtendedScryingContext(this, this.shared.nasSpellContext);
        await prepareMaleficiumContext(this, this.shared.nasSpellContext);
        await prepareEldritchResearcherContext(this, this.shared.nasSpellContext);
        await prepareSpellPerfectionContext(this, this.shared.nasSpellContext);
        await prepareSpontaneousMetafocusContext(this, this.shared.nasSpellContext);
        await prepareMagicalLineageContext(this, this.shared.nasSpellContext);
        await prepareWayangSpellhunterContext(this, this.shared.nasSpellContext);
        await prepareMaskFocusContext(this, this.shared.nasSpellContext);
        await applyMetamagicSelections(this, this.shared.nasSpellContext);
        applyEldritchResearcherPostMetamagic(this, this.shared.nasSpellContext);
        applyMaleficiumPostMetamagic(this, this.shared.nasSpellContext);
        if (this.shared?.rollData) {
          const intensifySourceParts = Array.isArray(this.shared.nasSpellContext?.damage?.parts)
            ? this.shared.nasSpellContext.damage.parts.map((part) => {
              if (Array.isArray(part)) return part[0] ?? "";
              if (part && typeof part === "object") return part.formula ?? "";
              return "";
            })
            : [];
          this.shared.rollData.nasMeta = {
            rollTransforms: this.shared.nasSpellContext?.rollPatch?.damage?.transforms ?? [],
            metamagicNames: this.shared.nasSpellContext?.metamagicNames ?? [],
            intensifySourceParts
          };
        }
        const restore = applyActionUseOverrides(this, this.shared.nasSpellContext);
        this.shared.nasRestoreOverrides = restore;
      } catch (_error) {}
      return result;
    },
    "MIXED"
  );

  libWrapper.register(
    MODULE.ID,
    "pf1.actionUse.ActionUse.prototype.executeScriptCalls",
    async function (wrapped, category = "use", ...args) {
      if (category !== "use") {
        return wrapped.apply(this, [category, ...args]);
      }

      if (shouldHandleAutomaticBuffs(this)) {
        await handleBuffAutomation(this);
      }
      await handleMirrorImageCast(this);
      if (this.shared?.reject === true || this.shared?.scriptData?.reject === true) {
        this.shared.scriptData ??= {};
        this.shared.scriptData.reject = true;
        return;
      }

      return wrapped.apply(this, [category, ...args]);
    },
    "MIXED"
  );

  libWrapper.register(
    MODULE.ID,
    "pf1.actionUse.ActionUse.prototype.getMessageData",
    async function (wrapped, ...args) {
      const result = await wrapped.apply(this, args);
      try {
        await resolveMirrorImagesForActionUse(this);
      } catch (_error) {}
      return result;
    },
    "MIXED"
  );

  Hooks.on("pf1PreDisplayActionUse", (actionUse) => {
    if (!shouldHandleMetamagic(actionUse)) return;
    const context = actionUse?.shared?.nasSpellContext;
    if (!context) return;
    try {
      actionUse.shared.chatData.flags ??= {};
      actionUse.shared.chatData.flags[MODULE.ID] ??= {};
      actionUse.shared.chatData.flags[MODULE.ID].actionOverrides ??= {};
      actionUse.shared.chatData.flags[MODULE.ID].actionOverrides.range = buildRangeOverride(actionUse);
      if (context.duration) {
        const durationOverride = {
          value: context.duration.value ?? "",
          units: context.duration.units ?? "",
          dismiss: Boolean(context.duration.dismiss),
          concentration: Boolean(context.duration.concentration)
        };
        actionUse.shared.chatData.flags[MODULE.ID].actionOverrides.duration = durationOverride;
        foundry.utils.setProperty(
          actionUse.shared,
          "templateData.item.system.actions.0.duration.value",
          durationOverride.value
        );
        foundry.utils.setProperty(
          actionUse.shared,
          "templateData.item.system.actions.0.duration.units",
          durationOverride.units
        );
        foundry.utils.setProperty(
          actionUse.shared,
          "templateData.item.system.actions.0.duration.dismiss",
          durationOverride.dismiss
        );
        foundry.utils.setProperty(
          actionUse.shared,
          "templateData.item.system.actions.0.duration.concentration",
          durationOverride.concentration
        );
        foundry.utils.setProperty(
          actionUse.shared,
          "chatData.system.actions.0.duration.value",
          durationOverride.value
        );
        foundry.utils.setProperty(
          actionUse.shared,
          "chatData.system.actions.0.duration.units",
          durationOverride.units
        );
      }

      const contextSaveDcNumeric = coerceNumericSaveDc(context?.save?.dc);
      if (contextSaveDcNumeric != null && contextHasSavingThrow(context, actionUse)) {
        const saveType = context?.save?.type ?? actionUse?.shared?.save ?? actionUse?.action?.save?.type ?? "";
        const saveTypeLabel = pf1?.config?.savingThrows?.[saveType] ?? saveType;
        const saveLabel = game.i18n.format("PF1.SavingThrowButtonLabel", {
          type: saveTypeLabel,
          dc: contextSaveDcNumeric.toString(),
        });
        const gmSensitiveLabel = game.i18n.format("PF1.SavingThrowButtonLabelGMSensitive", {
          save: saveTypeLabel,
        });

        actionUse.shared.saveDC = contextSaveDcNumeric;
        actionUse.shared.chatData.flags[MODULE.ID].actionOverrides.save = {
          dc: contextSaveDcNumeric,
          type: saveType,
        };
        foundry.utils.setProperty(actionUse.shared, "templateData.save.dc", contextSaveDcNumeric);
        foundry.utils.setProperty(actionUse.shared, "templateData.save.type", saveType);
        foundry.utils.setProperty(actionUse.shared, "templateData.save.label", saveLabel);
        foundry.utils.setProperty(actionUse.shared, "templateData.save.gmSensitiveLabel", gmSensitiveLabel);
        foundry.utils.setProperty(actionUse.shared, "templateData.item.system.save.dc", contextSaveDcNumeric);
        foundry.utils.setProperty(actionUse.shared, "chatData.system.save.dc", contextSaveDcNumeric);
        foundry.utils.setProperty(actionUse.shared, "chatData.system.save.type", saveType);
      }

      if (context.metamagic?.persistent || context.metamagic?.dazing) {
        actionUse.shared.chatData.flags[MODULE.ID].metamagic ??= {};
      }
      if (context.metamagic?.persistent) {
        actionUse.shared.chatData.flags[MODULE.ID].metamagic.persistent = true;
      }
      if (context.metamagic?.dazing) {
        actionUse.shared.chatData.flags[MODULE.ID].metamagic.dazing = true;
        actionUse.shared.chatData.flags[MODULE.ID].metamagic.dazingRounds = context.metamagic?.dazingRounds ?? 1;
        actionUse.shared.chatData.flags[MODULE.ID].metamagic.dazingSpellName =
          context.metamagic?.dazingSpellName ?? "";
      }

      if (Array.isArray(actionUse.shared.targets) && actionUse.shared.targets.length) {
        actionUse.shared.chatData.flags[MODULE.ID].targets = actionUse.shared.targets
          .map((target) => target?.document?.uuid ?? target?.uuid ?? target?.id)
          .filter(Boolean);
      }

      const activationOverride = context.activation;
      const extraFullRound = context.activationExtraFullRound;
      if (activationOverride || extraFullRound) {
        actionUse.shared.chatData.flags[MODULE.ID].actionOverrides.activation = {
          activation: activationOverride,
          extraFullRound: Boolean(extraFullRound)
        };
        if (activationOverride) {
          foundry.utils.setProperty(
            actionUse.shared,
            "templateData.item.system.actions.0.activation.type",
            activationOverride.type
          );
          foundry.utils.setProperty(
            actionUse.shared,
            "templateData.item.system.actions.0.activation.cost",
            activationOverride.cost
          );
          const unchainedType = activationOverride?.unchained?.type;
          const unchainedCost = activationOverride?.unchained?.cost;
          if (unchainedType !== undefined) {
            foundry.utils.setProperty(
              actionUse.shared,
              "templateData.item.system.actions.0.activation.unchained.type",
              unchainedType
            );
          }
          if (unchainedCost !== undefined) {
            foundry.utils.setProperty(
              actionUse.shared,
              "templateData.item.system.actions.0.activation.unchained.cost",
              unchainedCost
            );
          }
        }
        const originalActivation = actionUse?.action?.activation;
        const activationLabelPatch = updateChatInfoActivationLabel(
          actionUse?.shared?.templateData,
          originalActivation,
          activationOverride
        );
        const descriptionPatch = updateDescriptionActivationLabel(
          String(actionUse?.shared?.templateData?.description ?? ""),
          activationLabelPatch.beforeLabel,
          activationLabelPatch.afterLabel
        );
        if (descriptionPatch.changed) {
          actionUse.shared.templateData.description = descriptionPatch.html;
        }
      }

      const somaticLabel = game.i18n.localize(pf1?.config?.spellComponents?.somatic ?? "PF1.SpellComponents.Type.somatic.Abbr");
      const descriptionBeforeComponentsPatch = String(actionUse?.shared?.templateData?.description ?? "");

      const templateComponents = actionUse.shared?.templateData?.item?.system?.components;
      if (templateComponents && context.components) {
        Object.entries(context.components).forEach(([key, value]) => {
          foundry.utils.setProperty(actionUse.shared.templateData, `item.system.components.${key}`, value);
        });
      }

      const maskFocusComponentsPatch = patchMaskFocusComponentsInDescription(
        String(actionUse?.shared?.templateData?.description ?? ""),
        actionUse?.item ?? null,
        context
      );
      if (maskFocusComponentsPatch.changed) {
        actionUse.shared.templateData.description = maskFocusComponentsPatch.html;
      }

      const maskFocusDurationPatch = patchMaskFocusDurationInDescription(
        String(actionUse.shared?.templateData?.description ?? ""),
        context
      );
      if (maskFocusDurationPatch.changed) {
        actionUse.shared.templateData.description = maskFocusDurationPatch.html;
      }

      const isSilentSpellActive =
        context?.components?.verbal === false ||
        (Array.isArray(context?.metamagicNames) && context.metamagicNames.includes("Silent Spell"));
      const isStillSpellActive =
        context?.components?.somatic === false &&
        (
          (Array.isArray(context?.metamagic?.applied) && context.metamagic.applied.includes("Still Spell"))
          || (Array.isArray(context?.metamagicNames) && context.metamagicNames.includes("Still Spell"))
          || context?.metamagic?.oneBodyTwoMindsApplied === true
        );
      if ((isSilentSpellActive || isStillSpellActive) && typeof actionUse.shared?.templateData?.description === "string") {
        const beforeDescription = actionUse.shared.templateData.description;
        const nextDescription = sanitizeSpellDescriptionComponents(beforeDescription, {
          removeVerbal: isSilentSpellActive,
          removeSomatic: isStillSpellActive
        });
        actionUse.shared.templateData.description = nextDescription;
      }

      maybeApplyMetamagicChatCardNamePrefix(actionUse, context);
      applyMetamagicChatCardFooters(actionUse, context);
      const durationDisplay = buildChatDurationDisplay(context);
      const durationDescriptionPatch = patchDurationInDescription(
        String(actionUse.shared?.templateData?.description ?? ""),
        durationDisplay
      );
      if (durationDescriptionPatch.changed) {
        actionUse.shared.templateData.description = durationDescriptionPatch.html;
      }
      patchDurationInInfoProperties(actionUse.shared?.templateData?.properties, durationDisplay);
    } catch (_error) {}
  });

  libWrapper.register(
    MODULE.ID,
    "pf1.actionUse.ActionUse.prototype.addAttacks",
    async function (wrapped, ...args) {
      await wrapped.apply(this, args);

      if (!isGrappleSelected(this) || !this?.action?.hasAttack) return;

      const existingGrappleAttack = (this.shared?.attacks ?? []).find((attack) => isGrappleCmbAttack(attack));
      if (existingGrappleAttack?.chatAttack?.attack) return;

      const ChatAttackClass = this.shared?.chatAttacks?.[0]?.constructor;
      if (!ChatAttackClass) return;

      const rollData = this.shared?.rollData;
      if (!rollData) return;

      const syntheticAttack = existingGrappleAttack ?? {
        ...createGrappleCmbAttackEntry(),
        abstract: true,
        ammo: null,
        chargeCost: null,
        chatAttack: null,
      };
      if (!existingGrappleAttack) this.shared.attacks.push(syntheticAttack);

      const attackIndex = this.shared.attacks.indexOf(syntheticAttack);
      rollData.attackCount = attackIndex + (this.shared?.skipAttacks ?? 0);

      const chatAttack = new ChatAttackClass(this.action, {
        label: syntheticAttack.label,
        rollData,
        targets: game.user.targets,
        actionUse: this,
      });

      const conditionalParts = this._getConditionalParts(syntheticAttack, { index: attackIndex });
      await chatAttack.addAttack({
        extraParts: [...(this.shared?.attackBonus ?? []), syntheticAttack.attackBonus],
        conditionalParts,
      });

      syntheticAttack.chatAttack = chatAttack;
      this.shared.chatAttacks.push(chatAttack);
      delete rollData.attackCount;
    },
    "MIXED"
  );
}
