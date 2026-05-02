import { resolveEnglishName } from "./compendiumNameResolver.js";
import {
  getDictionaryNumber,
  getDictionaryPrefixedStrings,
  setDictionaryEntries,
  setDictionaryPrefixedStrings
} from "./itemDictionarySelection.js";
import { resolveMetamagicNameFromDatabase } from "../metamagic/metamagic.js";
import { MODULE } from "../../../common/module.js";
import { elementFromHtmlLike } from "../../../common/foundryCompat.js";
import { getKineticistMetamagicSources } from "../metamagic/classes/index.js";
import { getKineticInvocationSpellMetamagicSources } from "../metamagic/classes/index.js";
import { getHealersBlessingFeatureSources } from "../metamagic/classes/index.js";
import { getIntenseCelebrationFeatureSources } from "../metamagic/classes/index.js";
import { getNaniteBloodlineArcanaFeatureSources } from "../metamagic/classes/index.js";
import { getOneBodyTwoMindsFeatureSources } from "../metamagic/classes/index.js";
import { getPeerlessSpeedFeatureSources } from "../metamagic/classes/index.js";
import { getOracleSeekerFeatureSources } from "../metamagic/classes/index.js";
import { getOracleSuccorFinalRevelationFeatureSources } from "../metamagic/classes/index.js";
import {
  MIMIC_METAMAGIC_CHOICE_FLAG_PREFIX,
  getArcanistMetamagicFeatureSources,
  getMimicMetamagicFlagState,
  getPsychicMimicMetamagicFeatureSources,
  getPsychicMimicMetamagicState,
  MIMIC_METAMAGIC_FEATURE_ID,
  MIMIC_METAMAGIC_SELECTED_FLAG,
  getSorcererMetamagicFeatureSources,
  getWizardMetamagicFeatureSources
} from "../metamagic/classes/index.js";
import {
  getMagicalLineageSource,
  getMagicalLineageState,
  getRacialSpellLikeTraitSources,
  getTransmuterOfKoradaSource,
  getWayangSpellhunterSource,
  promptTransmuterOfKoradaSpellSelection,
  promptWayangSpellhunterSpellSelection,
  setTransmuterOfKoradaSelectedSpellName,
  setWayangSpellhunterSelectedSpellName,
  TRANSMUTER_OF_KORADA_ID,
  WAYANG_SPELLHUNTER_ID,
  setMagicalLineageState
} from "../metamagic/traits/index.js";
import {
  getMaleficiumSource,
  getMaleficiumState,
  spellItemHasEvilDescriptor,
  getEldritchResearcherSource,
  getEldritchResearcherSpellChoices,
  getEldritchResearcherState,
  getEldritchResearcherStateModes,
  getExtendedScryingSource,
  getExtendedScryingState,
  getSpontaneousMetafocusSource,
  getSpellPerfectionSource,
  getSpellPerfectionState,
  getMaskFocusSource,
  MASK_FOCUS_ID,
  promptEldritchResearcherSpellSelection,
  resolveSpellLabelFromChoices,
  setMaleficiumState,
  setEldritchResearcherState,
  setExtendedScryingState,
  setSpellPerfectionState
} from "../metamagic/feats/index.js";
import {
  METAMAGIC_DEFINITION as ExtendSpellDef,
  isDurationEligibleForExtendSpell
} from "../metamagic/extendSpell.js";
import { canIntensifyAnyDamagePart } from "../metamagic/intensifiedSpell.js";
import { createGrappleCmbAttackEntry } from "../conditions/grappled/grappled.js";

export const GRAPPLE_FORM_KEY = "grapple";
export const GRAPPLE_CMB_ATTACK_TYPE = "nas-grapple-cmb";
export const GRAPPLE_CMB_MARKER = "0[NAS_Grapple_CMB]";
export const METAMAGIC_FORM_KEY = "metamagic";
export const METAMAGIC_SELECT_KEY = "metamagicSelection";
export const METAMAGIC_NAMES_KEY = "metamagicNames";
export const METAMAGIC_DROPDOWN_KEY = "metamagicDropdownOpen";
export const METAMAGIC_OPTIONS_KEY = "metamagicOptions";
export const METAMAGIC_FEATURE_STATE_KEY = "classFeatures";
export const TRAIT_OPTIONS_KEY = "traitOptions";
export const FEAT_OPTIONS_KEY = "featOptions";
const METAMAGIC_PREVIEW_MODE_NONE = "none";
const METAMAGIC_PREVIEW_MODE_CONCISE = "concise";
const METAMAGIC_PREVIEW_MODE_DETAILED = "detailed";

const PERSISTENT_FEATURE_STATES = new Map();
const PERSISTENT_FEATURE_FLAG_KEY = "metamagic.persistentFeatureStates";
const PERSISTENT_FEATURE_IDS = new Set([
  "eldritchResearcher",
  "magicalLineage",
  "spellPerfection",
  "wayangSpellhunterMinata",
  "arcaneBloodline",
  "arcaneApotheosis",
  "grandMaestro",
  "seekerOfTheEternalEmperor",
  "healersBlessing",
  "intenseCelebration",
  "naniteBloodlineArcana"
]);
const maskFocusExtendCouplingRoots = new WeakSet();
const DIALOG_LAYOUT_FRAME_BY_APP_ID = new Map();
const METAMAGIC_SLOT_COSTS = {
  "Still Spell": 1,
  "Silent Spell": 1,
  "Enlarge Spell": 1,
  "Extend Spell": 1,
  "Reach Spell": 1,
  "Quicken Spell": 4,
  "Selective Spell": 1,
  "Dazing Spell": 3,
  "Persistent Spell": 2,
  "Intensified Spell": 1,
  "Maximize Spell": 3,
  "Empower Spell": 2,
};
const SPONTANEOUS_METAFOCUS_SELECTED_FLAG = "SMFSC";
const SPONTANEOUS_METAFOCUS_SPELL_FLAG_PREFIX = "SMFSS";

function localizeMetamagic(path) {
  return game.i18n.localize(`NAS.metamagic.${path}`);
}

function formatMetamagic(path, data = {}) {
  return game.i18n.format(`NAS.metamagic.${path}`, data);
}

class DialogStateTracker {
  static #trackedApplications = new Map();

  static get(appId, key) {
    return this.#trackedApplications.get(appId)?.get(key);
  }

  static set(appId, key, value) {
    if (!this.#trackedApplications.has(appId)) {
      this.#trackedApplications.set(appId, new Map());
    }
    this.#trackedApplications.get(appId)?.set(key, value);

    const toRemove = [...this.#trackedApplications.keys()].filter((id) => !ui.windows?.[id]);
    for (const id of toRemove) {
      this.#trackedApplications.delete(id);
    }
  }
}

function requestAttackDialogAutoLayout(dialog) {
  const appId = dialog?.appId;
  if (!appId || typeof dialog?.setPosition !== "function") return;

  const previousFrame = DIALOG_LAYOUT_FRAME_BY_APP_ID.get(appId);
  if (Number.isFinite(previousFrame)) {
    try {
      cancelAnimationFrame(previousFrame);
    } catch (_err) {}
  }

  try {
    dialog.setPosition({ height: "auto" });
  } catch (_err) {
    try {
      dialog.setPosition();
    } catch (_ignore) {}
  }

  const frameId = requestAnimationFrame(() => {
    DIALOG_LAYOUT_FRAME_BY_APP_ID.delete(appId);
    try {
      dialog.setPosition({ height: "auto" });
    } catch (_err) {
      try {
        dialog.setPosition();
      } catch (_ignore) {}
    }
  });

  DIALOG_LAYOUT_FRAME_BY_APP_ID.set(appId, frameId);
}

function getPersistentFeatureStateKey(actor, featureId) {
  const actorKey = actor?.uuid ?? actor?.id ?? "";
  const featureKey = (featureId ?? "").toString().trim();
  if (!actorKey || !featureKey) return "";
  return `${actorKey}::${featureKey}`;
}

function getPersistentFeatureState(actor, featureId) {
  const key = getPersistentFeatureStateKey(actor, featureId);
  if (!key) return undefined;
  const inMemory = PERSISTENT_FEATURE_STATES.get(key);
  if (inMemory !== undefined) return inMemory;
  const persisted = actor?.getFlag?.(MODULE.ID, PERSISTENT_FEATURE_FLAG_KEY);
  const persistedValue = persisted && typeof persisted === "object"
    ? persisted[(featureId ?? "").toString().trim()]
    : undefined;
  if (persistedValue === undefined) return undefined;
  const normalized = persistedValue === true;
  PERSISTENT_FEATURE_STATES.set(key, normalized);
  return normalized;
}

function setPersistentFeatureState(actor, featureId, enabled) {
  if (!PERSISTENT_FEATURE_IDS.has((featureId ?? "").toString())) return;
  const key = getPersistentFeatureStateKey(actor, featureId);
  if (!key) return;
  const normalized = Boolean(enabled);
  PERSISTENT_FEATURE_STATES.set(key, normalized);
  if (!actor || typeof actor?.setFlag !== "function") return;
  const persisted = actor?.getFlag?.(MODULE.ID, PERSISTENT_FEATURE_FLAG_KEY);
  const nextPersisted = persisted && typeof persisted === "object" ? { ...persisted } : {};
  nextPersisted[(featureId ?? "").toString().trim()] = normalized;
  void actor.setFlag(MODULE.ID, PERSISTENT_FEATURE_FLAG_KEY, nextPersisted);
}

function getDialogPersistentState(actor, featureId, fallback = false) {
  const value = getPersistentFeatureState(actor, featureId);
  return value === undefined ? fallback : value === true;
}

export function commitPersistentFeatureStatesFromOptions(actor, options) {
  const featureStates = getFeatureStatesFromOptions(options);
  if (!featureStates || typeof featureStates !== "object") return;
  Object.entries(featureStates).forEach(([featureId, enabled]) => {
    setPersistentFeatureState(actor, featureId, enabled === true);
  });
}

function getFlagsContainer(form) {
  let container = form.querySelector('div.form-group.stacked.flags');
  if (container) return container;

  container = document.createElement('div');
  container.classList.add('form-group', 'stacked', 'flags');
  const label = document.createElement('label');
  label.textContent = ` ${game.i18n.localize('PF1.Misc')} `;
  container.appendChild(label);
  const sibling = form.querySelector('.form-group.flags') || form.querySelector('.form-group');
  if (sibling) sibling.after(container);
  else form.appendChild(container);
  return container;
}

function createGrappleCmbDialogAttack(dialog) {
  const base = createGrappleCmbAttackEntry();
  const AttackUseAttackCtor = dialog?.attacks?.[0]?.constructor ?? pf1?.actionUse?.ActionUseAttack;
  const attack = AttackUseAttackCtor
    ? new AttackUseAttackCtor(base.label, GRAPPLE_CMB_MARKER, null, {
      abstract: true,
      type: GRAPPLE_CMB_ATTACK_TYPE,
    })
    : {
      label: base.label,
      attackBonus: GRAPPLE_CMB_MARKER,
      ammo: null,
      abstract: true,
      type: GRAPPLE_CMB_ATTACK_TYPE,
      chatAttack: null,
    };
  attack.attackBonusTotal = 0;
  return attack;
}

function setGrappleCmbAttackEnabled(dialog, enabled) {
  if (!dialog || !Array.isArray(dialog.attacks)) return false;

  const existingIndex = dialog.attacks.findIndex((attack) => attack?.type === GRAPPLE_CMB_ATTACK_TYPE);
  if (enabled) {
    if (existingIndex !== -1) return false;
    dialog.attacks.push(createGrappleCmbDialogAttack(dialog));
    return true;
  }

  if (existingIndex === -1) return false;
  if (typeof dialog.attacks.findSplice === "function") {
    dialog.attacks.findSplice((attack) => attack?.type === GRAPPLE_CMB_ATTACK_TYPE);
  } else {
    dialog.attacks.splice(existingIndex, 1);
  }
  return true;
}

function getActionItemType(dialog) {
  return dialog?.action?.item?.type ?? "";
}

function getActionActor(dialog) {
  return dialog?.action?.actor ?? dialog?.action?.item?.actor ?? dialog?.actor ?? null;
}

function isSpellLikeAbilityItem(dialog) {
  const item = dialog?.action?.item ?? null;
  const abilityType = item?.system?.abilityType ?? "";
  return (abilityType ?? "").toString().toLowerCase() === "sp";
}

function isMetamagicEligibleAction(dialog) {
  const type = getActionItemType(dialog);
  if (type === "spell") return true;
  return isSpellLikeAbilityItem(dialog);
}

function getSpellComponents(dialog) {
  return (
    dialog?.action?.components ??
    dialog?.action?.item?.system?.components ??
    {}
  );
}

function getSpellDuration(dialog) {
  return dialog?.action?.duration ?? dialog?.action?.item?.system?.duration ?? {};
}

function getSpellActivation(dialog) {
  return dialog?.action?.activation ?? dialog?.action?.item?.system?.activation ?? {};
}

function getSpellDamageParts(dialog) {
  return (
    dialog?.action?.damage?.parts ??
    dialog?.action?.item?.system?.damage?.parts ??
    []
  );
}

function hasDamageFormula(dialog) {
  const parts = getSpellDamageParts(dialog);
  if (!Array.isArray(parts) || !parts.length) return false;
  return parts.some((part) => {
    const formula = part?.formula ?? part?.[0];
    return typeof formula === "string" && formula.trim().length > 0;
  });
}

function getSpellCasterLevel(dialog) {
  const cl = dialog?.rollData?.cl;
  const value = Number(cl ?? 0);
  return Number.isFinite(value) ? value : 0;
}

function getSpellSaveType(dialog) {
  return (
    dialog?.action?.save?.type ??
    dialog?.action?.item?.system?.save?.type ??
    ""
  );
}

function getSpellRangeUnits(dialog) {
  const rawUnits =
    dialog?.action?.range?.units ??
    dialog?.action?.item?.system?.range?.units ??
    "";
  const normalized = rawUnits?.toString?.().toLowerCase() ?? "";
  if (normalized) return normalized;
  if (dialog?.action?.touch) return "touch";
  return "";
}

function getCanonicalMetamagicName(source) {
  const raw = source?.metaName ?? source?.label ?? "";
  return resolveMetamagicNameFromDatabase(raw) ?? raw;
}

function filteredSourcesIncludeExtendSpell(sources) {
  if (!Array.isArray(sources)) return false;
  return sources.some((source) => getCanonicalMetamagicName(source).toLowerCase() === "extend spell");
}

function canApplyExtendSpell(duration) {
  return isDurationEligibleForExtendSpell(duration);
}

function isTransmutationSchoolSpellForNanite(dialog) {
  const school = (dialog?.action?.item?.system?.school ?? "").toString().trim().toLowerCase();
  return school === "trs" || school === "transmutation" || school === "tra";
}

function getCurrentDialogTargetCount() {
  const targetSet = game?.user?.targets;
  if (typeof targetSet?.size === "number") return Math.max(0, targetSet.size);
  if (Array.isArray(targetSet)) return Math.max(0, targetSet.length);
  return 0;
}

function isNaniteBloodlineArcanaEligibleForDialog(dialog) {
  if (!isTransmutationSchoolSpellForNanite(dialog)) return false;
  if (!canApplyExtendSpell(getSpellDuration(dialog))) return false;
  return getCurrentDialogTargetCount() <= 1;
}

function isAreaEffectSpell(dialog) {
  const areaString = dialog?.action?.area ?? dialog?.action?.item?.system?.area ?? "";
  const templateType = dialog?.action?.measureTemplate?.type ?? dialog?.action?.item?.system?.measureTemplate?.type;
  return Boolean(areaString || templateType);
}

function isInstantDuration(duration) {
  const units = (duration?.units ?? "").toString().toLowerCase();
  return units === "inst" || units === "instantaneous";
}

function getSpellbookAbilityMod(dialog) {
  const mod = dialog?.rollData?.ablMod;
  const value = Number(mod ?? 0);
  return Number.isFinite(value) ? value : 0;
}

function getSpellBaseLevel(dialog) {
  const level = dialog?.action?.item?.system?.level ?? dialog?.item?.system?.level;
  const value = Number(level ?? 0);
  return Number.isFinite(value) ? value : 0;
}

function getDialogSaveDc(dialog) {
  if (typeof dialog?.action?.getDC === "function" && dialog?.rollData) {
    const dc = Number(dialog.action.getDC(dialog.rollData));
    if (Number.isFinite(dc)) return dc;
  }
  const actionDc = Number(dialog?.action?.save?.dc);
  if (Number.isFinite(actionDc)) return actionDc;
  const itemDc = Number(dialog?.action?.item?.system?.save?.dc);
  if (Number.isFinite(itemDc)) return itemDc;
  return null;
}

function canApplyQuickenSpell(activation) {
  const type = (activation?.type ?? "").toString().toLowerCase();
  const costValue = Number(activation?.cost ?? 1);
  const cost = Number.isFinite(costValue) ? costValue : 1;
  if (type === "swift") return false;
  if (!type) return true;
  if (["round", "full"].includes(type)) {
    return cost <= 1;
  }
  if (["minute", "hour", "day", "week", "month", "year"].includes(type)) return false;
  return true;
}

function canApplyEnlargeSpell(rangeUnits) {
  return ["close", "medium", "long"].includes(rangeUnits);
}

function canApplyIntensifiedSpell(dialog) {
  const cl = getSpellCasterLevel(dialog);
  if (!Number.isFinite(cl) || cl <= 0) return false;
  const parts = getSpellDamageParts(dialog);
  return canIntensifyAnyDamagePart(parts, cl);
}

function filterMetamagicSourcesForDialog(dialog, sources) {
  const components = getSpellComponents(dialog);
  const duration = getSpellDuration(dialog);
  const activation = getSpellActivation(dialog);
  const rangeUnits = getSpellRangeUnits(dialog);
  const isAreaEffect = isAreaEffectSpell(dialog);
  const abilityMod = getSpellbookAbilityMod(dialog);

  return sources.filter((source) => {
    const name = getCanonicalMetamagicName(source).toString().toLowerCase();
    if (name === "still spell") {
      return components?.somatic === true;
    }
    if (name === "silent spell") {
      return components?.verbal === true;
    }
    if (name === "extend spell") {
      return canApplyExtendSpell(duration);
    }
    if (name === "enlarge spell") {
      return canApplyEnlargeSpell(rangeUnits);
    }
    if (name === "reach spell") {
      return ["touch", "close", "medium"].includes(rangeUnits);
    }
    if (name === "quicken spell") {
      return canApplyQuickenSpell(activation);
    }
    if (name === "selective spell") {
      return isAreaEffect && isInstantDuration(duration) && abilityMod > 0;
    }
    if (name === "dazing spell") {
      return hasDamageFormula(dialog);
    }
    if (name === "persistent spell") {
      return Boolean(getSpellSaveType(dialog));
    }
    if (name === "heighten spell") {
      const baseLevel = getSpellBaseLevel(dialog);
      return baseLevel < 9;
    }
    if (name === "intensified spell") {
      return canApplyIntensifiedSpell(dialog);
    }
    if (name === "maximize spell") {
      return hasDamageFormula(dialog);
    }
    return true;
  });
}

function hasAvailableDailyUses(item) {
  const uses = item?.system?.uses;
  if (!uses) return true;
  if (uses?.per !== "day") return true;
  const value = Number(uses?.value ?? 0);
  return Number.isFinite(value) && value > 0;
}

function isMetamagicFeat(item) {
  if (item?.type !== "feat" || item?.subType !== "feat") return false;
  const tags = item?.system?.tags;
  if (!Array.isArray(tags)) return false;
  return tags.some((tag) => tag?.toString?.().toLowerCase().includes("metamagic"));
}

function isMetamagicRodName(name) {
  return Boolean(extractRodMetamagicPrefix(name));
}

function cleanMetamagicName(name) {
  return (name ?? "")
    .toString()
    .replace(/[,]/g, " ")
    .replace(/\b(lesser|greater)\b/gi, " ")
    .replace(/\bmetamagic\b/gi, " ")
    .replace(/\brod\b/gi, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function extractRodMetamagicPrefix(name) {
  const label = (name ?? "").toString();
  if (!label) return "";
  const lower = label.toLowerCase();
  if (!lower.includes("rod") || !lower.includes("metamagic")) return "";

  const ofMatch = label.match(/rod\s+of\s+(.+?)\s+metamagic/i);
  if (ofMatch?.[1]) return cleanMetamagicName(ofMatch[1]);

  const rodMatch = label.match(/(.+?)\s+metamagic\s+rod/i);
  if (rodMatch?.[1]) return cleanMetamagicName(rodMatch[1]);

  const suffixMatch = label.match(/metamagic\s+rod[, ]+\s*(.+)$/i);
  if (suffixMatch?.[1]) return cleanMetamagicName(suffixMatch[1]);

  return "";
}

async function getItemEnglishName(item) {
  const itemName = (item?.name ?? "").toString().trim();
  if (!itemName) return itemName;

  const databaseName = resolveMetamagicNameFromDatabase(itemName);
  if (databaseName) return databaseName;

  return resolveEnglishName(itemName, { documentName: "Item", deepScanMode: "off" });
}

async function getAvailableMetamagicSources(actor, options = {}) {
  if (!actor?.items?.size) return [];
  const { resolveEnglishNames = false } = options;
  const items = Array.from(actor.items).filter(
    (item) =>
      (item?.type === "feat" && item?.subType === "feat") ||
      item?.constructor?.name === "ItemEquipmentPF"
  );
  items.sort((a, b) => {
    const aHasOriginal = Boolean(a?.flags?.babele?.originalName);
    const bHasOriginal = Boolean(b?.flags?.babele?.originalName);
    return Number(bHasOriginal) - Number(aHasOriginal);
  });
  const resolved = await Promise.all(
    items.map(async (item) => {
      const originalName = item?.flags?.babele?.originalName;
      const englishName = originalName
        ? originalName
        : resolveEnglishNames
          ? await getItemEnglishName(item)
          : item?.name ?? "";

      return { item, englishName };
    })
  );

  return resolved
    .map(({ item, englishName }) => {
      if (!item || !englishName) return null;

      if (isMetamagicFeat(item)) {
        return {
          item,
          type: "feat",
          label: item?.name ?? cleanMetamagicName(englishName),
          metaName: englishName,
        };
      }

      const rodPrefix =
        extractRodMetamagicPrefix(englishName) ||
        extractRodMetamagicPrefix(item?.flags?.babele?.originalName) ||
        extractRodMetamagicPrefix(item?.name);
      if (!rodPrefix) return null;
      if (!hasAvailableDailyUses(item)) return null;

      return {
        item,
        type: "rod",
        label: item?.name ?? rodPrefix,
        metaName: englishName,
      };
    })
    .filter(Boolean);
}

export function addGrappleCheckbox(dialog, html) {
  if (!(dialog instanceof pf1.applications.AttackDialog)) return;
  if (!dialog.action?.hasAttack) return;

  const root = elementFromHtmlLike(html);
  if (!root) return;
  const form = root.querySelector?.('form') ?? root;
  if (!form) return;

  if (form.querySelector(`input[name="${GRAPPLE_FORM_KEY}"]`)) return;

  const labelText = game.i18n.localize('NAS.conditions.main.GrappleCheckbox');
  const container = getFlagsContainer(form);
  const metamagicDropdown = container.querySelector("details.metamagic-dropdown");

  const labelElement = document.createElement('label');
  labelElement.classList.add('checkbox');
  const input = document.createElement('input');
  input.type = 'checkbox';
  input.name = GRAPPLE_FORM_KEY;
  const storedValue = DialogStateTracker.get(dialog.appId, GRAPPLE_FORM_KEY);
  if (storedValue !== undefined) {
    input.checked = storedValue;
  }
  input.addEventListener('change', () => {
    DialogStateTracker.set(dialog.appId, GRAPPLE_FORM_KEY, input.checked);
    const didChangeRows = setGrappleCmbAttackEnabled(dialog, input.checked === true);
    if (didChangeRows) {
      dialog.render();
      return;
    }
    requestAttackDialogAutoLayout(dialog);
  });
  labelElement.textContent = ` ${labelText} `;
  labelElement.insertBefore(input, labelElement.firstChild);
  container.appendChild(labelElement);

  const didSyncRows = setGrappleCmbAttackEnabled(dialog, input.checked === true);
  if (didSyncRows) {
    dialog.render();
    return;
  }
  requestAttackDialogAutoLayout(dialog);
}

export async function addMetamagicCheckbox(dialog, html) {
  if (!(dialog instanceof pf1.applications.AttackDialog)) return;
  if (!isMetamagicEligibleAction(dialog)) return;

  const actor = getActionActor(dialog);
  const item = dialog?.action?.item ?? null;
  const root = elementFromHtmlLike(html);
  if (!root) return;
  const form = root.querySelector?.('form') ?? root;
  if (!form) return;

  let metamagicSources = [];
  let classFeatureSources = [];
  let healerBlessingStandaloneSources = [];
  let intenseCelebrationStandaloneSources = [];
  let naniteBloodlineArcanaStandaloneSources = [];
  let oneBodyTwoMindsStandaloneSources = [];
  let peerlessSpeedStandaloneSources = [];
  let succorFinalRevelationStandaloneSources = [];
  let mimicMetamagicStandaloneSources = [];
  let traitModifierSources = [];
  let eldritchResearcherSource = null;
  let spellPerfectionSource = null;
  let spontaneousMetafocusSource = null;
  let magicalLineageSource = null;
  let extendedScryingSource = null;
  let maleficiumSource = null;
  let maskFocusSource = null;
  let checkboxLabel = "Metamagic";
  if (getActionItemType(dialog) === "spell") {
    metamagicSources = await getAvailableMetamagicSources(actor, { resolveEnglishNames: false });
    const invocationSources = await getKineticInvocationSpellMetamagicSources(actor, item);
    if (invocationSources?.length) {
      metamagicSources = [...metamagicSources, ...invocationSources];
      checkboxLabel = "Invocation";
    }
    const [
      sorcererFeatureSources,
      oracleFeatureSources,
      oracleSuccorFeatureSources,
      wizardFeatureSources,
      arcanistFeatureSources,
      healersBlessingSources,
      intenseCelebrationSources,
      naniteBloodlineArcanaSources,
      oneBodyTwoMindsSources,
      peerlessSpeedSources,
      mimicMetamagicSources
    ] = await Promise.all([
      getSorcererMetamagicFeatureSources(actor, item),
      getOracleSeekerFeatureSources(actor, item),
      getOracleSuccorFinalRevelationFeatureSources(actor, item),
      getWizardMetamagicFeatureSources(actor, item),
      getArcanistMetamagicFeatureSources(actor, item),
      getHealersBlessingFeatureSources(actor, item),
      getIntenseCelebrationFeatureSources(actor, item),
      getNaniteBloodlineArcanaFeatureSources(actor, item, {
        durationOverride: dialog?.action?.duration ?? null
      }),
      getOneBodyTwoMindsFeatureSources(actor, item),
      getPeerlessSpeedFeatureSources(actor, item),
      getPsychicMimicMetamagicFeatureSources(actor, item)
    ]);
    classFeatureSources = [
      ...(Array.isArray(sorcererFeatureSources) ? sorcererFeatureSources : []),
      ...(Array.isArray(oracleFeatureSources) ? oracleFeatureSources : []),
      ...(Array.isArray(wizardFeatureSources) ? wizardFeatureSources : []),
      ...(Array.isArray(arcanistFeatureSources) ? arcanistFeatureSources : [])
    ];
    succorFinalRevelationStandaloneSources = Array.isArray(oracleSuccorFeatureSources) ? oracleSuccorFeatureSources : [];
    healerBlessingStandaloneSources = Array.isArray(healersBlessingSources) ? healersBlessingSources : [];
    intenseCelebrationStandaloneSources = Array.isArray(intenseCelebrationSources) ? intenseCelebrationSources : [];
    naniteBloodlineArcanaStandaloneSources = Array.isArray(naniteBloodlineArcanaSources) ? naniteBloodlineArcanaSources : [];
    oneBodyTwoMindsStandaloneSources = Array.isArray(oneBodyTwoMindsSources) ? oneBodyTwoMindsSources : [];
    peerlessSpeedStandaloneSources = Array.isArray(peerlessSpeedSources) ? peerlessSpeedSources : [];
    mimicMetamagicStandaloneSources = Array.isArray(mimicMetamagicSources) ? mimicMetamagicSources : [];
    traitModifierSources = await getRacialSpellLikeTraitSources(actor, item);
    const transmuterSource = await getTransmuterOfKoradaSource(actor, item);
    const wayangSpellhunterSource = await getWayangSpellhunterSource(actor, item);
    const shouldIncludeTransmuterSource = Boolean(transmuterSource) && (
      transmuterSource?.hasRemaining === true &&
      (transmuterSource?.requiresSpellSelection === true || transmuterSource?.selectedSpellMatches === true)
    );
    if (shouldIncludeTransmuterSource) traitModifierSources.push(transmuterSource);
    if (wayangSpellhunterSource) traitModifierSources.push(wayangSpellhunterSource);
    eldritchResearcherSource = await getEldritchResearcherSource(actor);
    spellPerfectionSource = await getSpellPerfectionSource(actor);
    spontaneousMetafocusSource = await getSpontaneousMetafocusSource(actor);
    magicalLineageSource = await getMagicalLineageSource(actor);
    const candidateExtendedScryingSource = await getExtendedScryingSource(actor);
    const rawSubschool = item?.system?.subschool;
    const normalizedSubschool = extractSubschoolValues(rawSubschool)
      .map((v) => (v ?? "").toString().trim().toLowerCase());
    const hasScryingSubschool = normalizedSubschool.includes("scrying");
    extendedScryingSource = hasScryingSubschool ? candidateExtendedScryingSource : null;
    maleficiumSource = isEvilDescriptorSpellItem(item) ? await getMaleficiumSource(actor) : null;
    maskFocusSource = !isSpellLikeSpellbookItem(item) ? await getMaskFocusSource(actor) : null;
  } else if (isSpellLikeAbilityItem(dialog)) {
    metamagicSources = await getKineticistMetamagicSources(actor, item);
    checkboxLabel = "Metakinesis";
    traitModifierSources = await getRacialSpellLikeTraitSources(actor, item);
  }
  const filteredSources = filterMetamagicSourcesForDialog(dialog, metamagicSources);
  const maskFocusForMetamagicUi =
    maskFocusSource && filteredSourcesIncludeExtendSpell(filteredSources) ? maskFocusSource : null;
  if (filteredSources.length) {
    renderMetamagicControls(dialog, form, filteredSources, {
      actor,
      actionItem: item,
      checkboxLabel,
      classFeatureSources,
      maskFocusSource: maskFocusForMetamagicUi,
      baseSpellLevel: getSpellBaseLevel(dialog),
      baseSaveDc: getDialogSaveDc(dialog),
    });
  }
  if (traitModifierSources.length) {
    renderTraitModifierControls(dialog, form, traitModifierSources);
  }
  if (healerBlessingStandaloneSources.length) {
    renderHealersBlessingStandaloneControl(dialog, form, actor, healerBlessingStandaloneSources);
  }
  if (intenseCelebrationStandaloneSources.length) {
    renderIntenseCelebrationStandaloneControl(dialog, form, actor, intenseCelebrationStandaloneSources);
  }
  if (naniteBloodlineArcanaStandaloneSources.length) {
    renderNaniteBloodlineArcanaStandaloneControl(dialog, form, actor, naniteBloodlineArcanaStandaloneSources);
  }
  if (oneBodyTwoMindsStandaloneSources.length) {
    renderOneBodyTwoMindsStandaloneControl(dialog, form, actor, oneBodyTwoMindsStandaloneSources);
  }
  if (peerlessSpeedStandaloneSources.length) {
    renderPeerlessSpeedStandaloneControl(dialog, form, actor, peerlessSpeedStandaloneSources);
  }
  if (succorFinalRevelationStandaloneSources.length) {
    renderSuccorFinalRevelationStandaloneControl(dialog, form, actor, succorFinalRevelationStandaloneSources);
  }
  if (mimicMetamagicStandaloneSources.length) {
    renderMimicMetamagicStandaloneControl(dialog, form, actor, mimicMetamagicStandaloneSources);
  }
  if (spellPerfectionSource) {
    await renderSpellPerfectionStandaloneControl(dialog, form, actor, spellPerfectionSource);
  }
  if (spontaneousMetafocusSource) {
    await renderSpontaneousMetafocusStandaloneControl(dialog, form, actor, spontaneousMetafocusSource);
  }
  if (eldritchResearcherSource || magicalLineageSource || extendedScryingSource || maleficiumSource) {
    await renderEldritchResearcherControls(
      dialog,
      form,
      actor,
      eldritchResearcherSource,
      extendedScryingSource,
      magicalLineageSource,
      maleficiumSource
    );
  }

  const isEnglish = (game?.i18n?.lang ?? "en").toLowerCase().startsWith("en");
  const canUseBabele = game?.modules?.get("babele")?.active;
  const shouldResolveEnglishName = !isEnglish && canUseBabele;
  if (shouldResolveEnglishName && getActionItemType(dialog) === "spell") {
    void refreshMetamagicControls(dialog, form, actor, classFeatureSources);
  }
}

function renderMetamagicControls(dialog, form, metamagicSources, options = {}) {
  if (form.querySelector(`input[name="${METAMAGIC_FORM_KEY}"]`)) return;

  const container = getFlagsContainer(form);
  const labelElement = document.createElement('label');
  labelElement.classList.add('checkbox');
  const input = document.createElement('input');
  input.type = 'checkbox';
  input.name = METAMAGIC_FORM_KEY;
  const storedChecked = DialogStateTracker.get(dialog.appId, METAMAGIC_FORM_KEY);
  if (storedChecked !== undefined) {
    input.checked = storedChecked;
  }
  labelElement.textContent = ` ${(options.checkboxLabel ?? 'Metamagic').toString()} `;
  labelElement.insertBefore(input, labelElement.firstChild);
  container.appendChild(labelElement);

  const dataInput = document.createElement('input');
  dataInput.type = 'hidden';
  dataInput.name = METAMAGIC_NAMES_KEY;
  container.appendChild(dataInput);

  const optionsInput = document.createElement('input');
  optionsInput.type = 'hidden';
  optionsInput.name = METAMAGIC_OPTIONS_KEY;
  container.appendChild(optionsInput);

  const dropdown = document.createElement('details');
  dropdown.classList.add('metamagic-dropdown');
  dropdown.style.display = 'none';
  const summary = document.createElement('summary');
  summary.textContent = localizeMetamagic("dialogs.selectMetamagic");
  dropdown.appendChild(summary);
  const listContainer = document.createElement('div');
  listContainer.classList.add('metamagic-options');
  listContainer.style.maxHeight = '160px';
  listContainer.style.overflowY = 'auto';
  listContainer.style.display = 'grid';
  listContainer.style.gridTemplateColumns = 'repeat(3, minmax(0, 1fr))';
  listContainer.style.columnGap = '8px';

  const infoContainer = document.createElement('div');
  infoContainer.classList.add('nas-metamagic-info');
  infoContainer.style.border = '1px solid var(--color-border-light-tertiary, #888)';
  infoContainer.style.borderRadius = '4px';
  infoContainer.style.padding = '6px 8px';
  infoContainer.style.marginBottom = '8px';
  infoContainer.style.fontSize = '12px';
  infoContainer.style.lineHeight = '1.35';
  infoContainer.style.overflowWrap = 'break-word';
  dropdown.appendChild(infoContainer);

  const featureContainer = document.createElement('div');
  featureContainer.classList.add('nas-metamagic-features');
  featureContainer.style.border = '1px solid var(--color-border-light-tertiary, #888)';
  featureContainer.style.borderRadius = '4px';
  featureContainer.style.padding = '6px 8px';
  featureContainer.style.marginBottom = '8px';
  featureContainer.style.fontSize = '12px';
  featureContainer.style.lineHeight = '1.35';
  featureContainer.style.display = 'none';
  featureContainer.style.overflowWrap = 'break-word';
  dropdown.appendChild(featureContainer);

  dropdown.appendChild(listContainer);
  container.appendChild(dropdown);

  if (options.maskFocusSource) {
    ensureFeatOptionsInput(form, dialog);
  }
  updateMetamagicCheckboxOptions(dialog, listContainer, dataInput, optionsInput, metamagicSources, {
    actor: options.actor ?? null,
    form,
    actionItem: options.actionItem ?? null,
    baseSpellLevel: Number(options.baseSpellLevel ?? 0),
    baseSaveDc: Number.isFinite(Number(options.baseSaveDc)) ? Number(options.baseSaveDc) : null,
    featureContainer,
    featureSources: Array.isArray(options.classFeatureSources) ? options.classFeatureSources : [],
    maskFocusSource: options.maskFocusSource ?? null,
    infoContainer,
    metamagicToggleInput: input
  });
  dropdown.style.display = input.checked ? '' : 'none';
  const storedOpen = DialogStateTracker.get(dialog.appId, METAMAGIC_DROPDOWN_KEY);
  if (storedOpen !== undefined) {
    dropdown.open = storedOpen;
  }

  input.addEventListener('change', () => {
    DialogStateTracker.set(dialog.appId, METAMAGIC_FORM_KEY, input.checked);
    dropdown.style.display = input.checked ? '' : 'none';
    updateMetamagicNames(dataInput, input.checked, listContainer);
    updateMetamagicInfoSection(infoContainer, {
      baseSpellLevel: Number(options.baseSpellLevel ?? 0),
      baseSaveDc: Number.isFinite(Number(options.baseSaveDc)) ? Number(options.baseSaveDc) : null,
      metamagicEnabled: input.checked,
      options: DialogStateTracker.get(dialog.appId, METAMAGIC_OPTIONS_KEY) || {},
      featOptions: getCurrentFeatOptions(dialog, form),
      actionItem: options.actionItem ?? null,
      selectedMetaNames: getSelectedMetaNames(listContainer),
      dialog
    });
    requestAttackDialogAutoLayout(dialog);
  });
  dropdown.addEventListener('toggle', () => {
    DialogStateTracker.set(dialog.appId, METAMAGIC_DROPDOWN_KEY, dropdown.open);
    requestAttackDialogAutoLayout(dialog);
  });
  form.addEventListener('submit', () => {
    const currentOptions = DialogStateTracker.get(dialog.appId, METAMAGIC_OPTIONS_KEY) || {};
    commitPersistentFeatureStatesFromOptions(options.actor ?? null, currentOptions);
  }, { once: true });

  requestAttackDialogAutoLayout(dialog);
}

function updateMetamagicNames(dataInput, isChecked, listContainer) {
  if (!isChecked) {
    dataInput.value = "";
    return;
  }

  const selections = Array.from(
    listContainer.querySelectorAll('input[type="checkbox"][data-meta-name]:checked')
  )
    .map((input) => input.dataset.metaName)
    .filter(Boolean);

  dataInput.value = selections.length ? JSON.stringify(selections) : "";
}

function updateMetamagicOptionsInput(optionsInput, options) {
  const hasOptions = options && Object.keys(options).length > 0;
  optionsInput.value = hasOptions ? JSON.stringify(options) : "";
}

function parseMetamagicOptionsInput(raw) {
  if (!raw) return {};
  try {
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch {
    return {};
  }
}

function parseFeatOptionsInput(raw) {
  if (!raw) return {};
  try {
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch {
    return {};
  }
}

function readFeatOptionsPayload(dialog, form) {
  const tracked = DialogStateTracker.get(dialog.appId, FEAT_OPTIONS_KEY);
  if (tracked && typeof tracked === "object") return { ...tracked };
  const inp = form?.querySelector?.(`input[name="${FEAT_OPTIONS_KEY}"]`);
  return parseFeatOptionsInput(inp?.value ?? "");
}

function writeFeatOptionsPayload(dialog, form, payload) {
  if (!dialog || !form) return;
  DialogStateTracker.set(dialog.appId, FEAT_OPTIONS_KEY, payload);
  const inp = form.querySelector(`input[name="${FEAT_OPTIONS_KEY}"]`);
  if (inp) inp.value = JSON.stringify(payload);
}

function ensureFeatOptionsInput(form, dialog) {
  if (!form || !dialog) return null;
  let inp = form.querySelector(`input[name="${FEAT_OPTIONS_KEY}"]`);
  if (inp) return inp;
  const container = getFlagsContainer(form);
  const metamagicDropdown = container.querySelector("details.metamagic-dropdown");
  inp = document.createElement("input");
  inp.type = "hidden";
  inp.name = FEAT_OPTIONS_KEY;
  if (metamagicDropdown?.parentElement === container) {
    container.insertBefore(inp, metamagicDropdown);
  } else {
    container.appendChild(inp);
  }
  const tracked = DialogStateTracker.get(dialog.appId, FEAT_OPTIONS_KEY);
  if (tracked && typeof tracked === "object") {
    inp.value = JSON.stringify(tracked);
  }
  return inp;
}

function patchFeatOptionsMaskFocus(dialog, form, enabled) {
  const payload = readFeatOptionsPayload(dialog, form);
  payload[MASK_FOCUS_ID] = { enabled: enabled === true };
  writeFeatOptionsPayload(dialog, form, payload);
}

function wireMaskFocusExtendCoupling(form, dialog, listContainer) {
  if (!form || !dialog || !listContainer || maskFocusExtendCouplingRoots.has(listContainer)) return;
  maskFocusExtendCouplingRoots.add(listContainer);
  listContainer.addEventListener("change", (ev) => {
    const t = ev.target;
    if (t?.dataset?.metaName !== ExtendSpellDef.name) return;
    if (t.checked !== false) return;
    const maskInput = form.querySelector("label.nas-mask-focus input[type=\"checkbox\"]");
    if (!maskInput?.checked) return;
    maskInput.checked = false;
    patchFeatOptionsMaskFocus(dialog, form, false);
    const infoContainer = form.querySelector(".nas-metamagic-info");
    const metamagicToggleInput = form.querySelector(`input[name="${METAMAGIC_FORM_KEY}"]`);
    if (infoContainer) {
      updateMetamagicInfoSection(infoContainer, {
        baseSpellLevel: getSpellBaseLevel(dialog),
        baseSaveDc: getDialogSaveDc(dialog),
        metamagicEnabled: metamagicToggleInput?.checked !== false,
        options: getCurrentMetamagicOptions(dialog, form),
        featOptions: readFeatOptionsPayload(dialog, form),
        actionItem: dialog?.action?.item ?? null,
        selectedMetaNames: getSelectedMetaNames(listContainer),
        dialog
      });
    }
  });
}

function ensureMetamagicOptionsInput(form) {
  const existing = form.querySelector(`input[name="${METAMAGIC_OPTIONS_KEY}"]`);
  if (existing) return existing;
  const container = getFlagsContainer(form);
  const created = document.createElement("input");
  created.type = "hidden";
  created.name = METAMAGIC_OPTIONS_KEY;
  container.appendChild(created);
  return created;
}

function getCurrentMetamagicOptions(dialog, form) {
  const tracked = DialogStateTracker.get(dialog.appId, METAMAGIC_OPTIONS_KEY);
  if (tracked && typeof tracked === "object") return tracked;
  const optionsInput = form.querySelector(`input[name="${METAMAGIC_OPTIONS_KEY}"]`);
  return parseMetamagicOptionsInput(optionsInput?.value ?? "");
}

function getCurrentFeatOptions(dialog, form) {
  const tracked = DialogStateTracker.get(dialog.appId, FEAT_OPTIONS_KEY);
  if (tracked && typeof tracked === "object") return tracked;
  if (!form) return {};
  const optionsInput = form.querySelector(`input[name="${FEAT_OPTIONS_KEY}"]`);
  return parseFeatOptionsInput(optionsInput?.value ?? "");
}

function getSelectedMetaNames(listContainer) {
  return Array.from(
    listContainer.querySelectorAll('input[type="checkbox"][data-meta-name]:checked')
  )
    .map((input) => (input.dataset.metaName ?? "").toString())
    .filter(Boolean);
}

function getFeatureStatesFromOptions(options) {
  const raw = options?.[METAMAGIC_FEATURE_STATE_KEY];
  if (!raw || typeof raw !== "object") return {};
  return raw;
}

function setFeatureStatesInOptions(options, featureStates) {
  const nextOptions = { ...(options ?? {}) };
  const hasAny = featureStates && Object.keys(featureStates).length > 0;
  if (hasAny) {
    nextOptions[METAMAGIC_FEATURE_STATE_KEY] = featureStates;
  } else {
    delete nextOptions[METAMAGIC_FEATURE_STATE_KEY];
  }
  return nextOptions;
}

function parseTraitSelectionInput(rawValue) {
  if (!rawValue) return [];
  try {
    const parsed = JSON.parse(rawValue);
    return Array.isArray(parsed) ? parsed.map((value) => `${value}`).filter(Boolean) : [];
  } catch {
    return [];
  }
}

function isSpellLikeSpellbookItem(item) {
  return (item?.system?.spellbook ?? "").toString().trim().toLowerCase() === "spelllike";
}

function extractSubschoolValues(rawSubschool) {
  if (Array.isArray(rawSubschool)) return rawSubschool;
  if (typeof rawSubschool === "string") return [rawSubschool];
  if (rawSubschool && typeof rawSubschool === "object") {
    const values = [];
    if (Array.isArray(rawSubschool.base)) values.push(...rawSubschool.base);
    if (Array.isArray(rawSubschool.names)) values.push(...rawSubschool.names);
    if (Array.isArray(rawSubschool.total)) values.push(...rawSubschool.total);
    return values;
  }
  return [];
}

async function renderEldritchResearcherControls(
  dialog,
  form,
  actor,
  source,
  extendedScryingSource = null,
  magicalLineageSource = null,
  maleficiumSource = null
) {
  if (!form || !actor) return;
  if (!source && !extendedScryingSource && !magicalLineageSource && !maleficiumSource) return;
  form.querySelectorAll("label.nas-eldritch-researcher").forEach((node) => node.remove());
  form.querySelectorAll("label.nas-magical-lineage").forEach((node) => node.remove());
  form.querySelectorAll("label.nas-extended-scrying").forEach((node) => node.remove());
  form.querySelectorAll("label.nas-maleficium").forEach((node) => node.remove());
  const existingInput = form.querySelector(`input[name="${FEAT_OPTIONS_KEY}"]`);
  let priorFeatPayload = existingInput ? parseFeatOptionsInput(existingInput.value) : {};
  if (!Object.keys(priorFeatPayload).length) {
    const trackedFeat = DialogStateTracker.get(dialog.appId, FEAT_OPTIONS_KEY);
    if (trackedFeat && typeof trackedFeat === "object") {
      priorFeatPayload = trackedFeat;
    }
  }
  if (existingInput) existingInput.remove();

  const container = getFlagsContainer(form);
  const metamagicDropdown = container.querySelector("details.metamagic-dropdown");
  const featureContainer = form.querySelector(".nas-metamagic-features");
  const optionsInput = document.createElement("input");
  optionsInput.type = "hidden";
  optionsInput.name = FEAT_OPTIONS_KEY;
  if (metamagicDropdown?.parentElement === container) {
    container.insertBefore(optionsInput, metamagicDropdown);
  } else {
    container.appendChild(optionsInput);
  }

  const modes = getEldritchResearcherStateModes();
  const spellChoices = getEldritchResearcherSpellChoices(actor);
  const storedState = await getEldritchResearcherState(actor);
  const persistentEldritch = getDialogPersistentState(actor, source?.id, storedState.enabled === true);
  const state = {
    enabled: persistentEldritch,
    mode: storedState.mode === modes.complete ? modes.complete : modes.incomplete,
    selectedSpellUuid: (storedState.selectedSpellUuid ?? "").toString(),
    selectedSpellLabel: (storedState.selectedSpellLabel ?? "").toString(),
    selectedSpellName: (storedState.selectedSpellName ?? storedState.selectedSpellLabel ?? "").toString()
  };
  if (!state.selectedSpellUuid && state.selectedSpellName) {
    const byName = spellChoices.find(
      (entry) => (entry?.label ?? "").toString().trim().toLowerCase() === state.selectedSpellName.trim().toLowerCase()
    ) ?? null;
    state.selectedSpellUuid = byName?.uuid || byName?.id || "";
    state.selectedSpellLabel = byName?.label ?? state.selectedSpellName;
  }
  if (!state.selectedSpellLabel && state.selectedSpellUuid) {
    state.selectedSpellLabel = resolveSpellLabelFromChoices(spellChoices, state.selectedSpellUuid, "");
  }
  const storedMagicalLineageState = await getMagicalLineageState(actor);
  const persistentMagicalLineage = getDialogPersistentState(
    actor,
    magicalLineageSource?.id,
    storedMagicalLineageState.enabled === true
  );
  const magicalLineageState = {
    enabled: persistentMagicalLineage,
    selectedSpellUuid: (storedMagicalLineageState.selectedSpellUuid ?? "").toString(),
    selectedSpellLabel: (storedMagicalLineageState.selectedSpellLabel ?? "").toString(),
    selectedSpellName: (storedMagicalLineageState.selectedSpellName ?? storedMagicalLineageState.selectedSpellLabel ?? "").toString()
  };
  if (!magicalLineageState.selectedSpellUuid && magicalLineageState.selectedSpellName) {
    const byName = spellChoices.find(
      (entry) => (entry?.label ?? "").toString().trim().toLowerCase() === magicalLineageState.selectedSpellName.trim().toLowerCase()
    ) ?? null;
    magicalLineageState.selectedSpellUuid = byName?.uuid || byName?.id || "";
    magicalLineageState.selectedSpellLabel = byName?.label ?? magicalLineageState.selectedSpellName;
  }
  if (!magicalLineageState.selectedSpellLabel && magicalLineageState.selectedSpellUuid) {
    magicalLineageState.selectedSpellLabel = resolveSpellLabelFromChoices(
      spellChoices,
      magicalLineageState.selectedSpellUuid,
      ""
    );
  }
  const extendedScryingState = {
    enabled: getExtendedScryingState(actor) === true
  };
  const maleficiumState = {
    enabled: getMaleficiumState(actor) === true,
    damnationCount: Number(maleficiumSource?.damnationCount ?? 0)
  };

  let label = null;
  let input = null;
  let modeButton = null;
  let magicalLineageLabel = null;
  let magicalLineageInput = null;
  let maleficiumLabel = null;
  let maleficiumInput = null;

  if (source) {
    label = document.createElement("label");
    label.classList.add("checkbox");
    label.classList.add("nas-eldritch-researcher");
    input = document.createElement("input");
    input.type = "checkbox";
    input.checked = state.enabled;
    label.textContent = ` ${source.label} `;
    label.insertBefore(input, label.firstChild);

    modeButton = document.createElement("button");
    modeButton.type = "button";
    modeButton.classList.add("button");
    modeButton.style.padding = "0 6px";
    modeButton.style.height = "24px";
    modeButton.style.lineHeight = "20px";
    modeButton.style.minWidth = "0";
    modeButton.style.width = "auto";
    modeButton.style.flex = "0 0 auto";
    modeButton.style.display = "inline-block";
    modeButton.style.verticalAlign = "middle";
    modeButton.style.textAlign = "center";
    label.appendChild(modeButton);
  }

  if (magicalLineageSource) {
    magicalLineageLabel = document.createElement("label");
    magicalLineageLabel.classList.add("checkbox");
    magicalLineageLabel.classList.add("nas-magical-lineage");
    magicalLineageInput = document.createElement("input");
    magicalLineageInput.type = "checkbox";
    magicalLineageInput.checked = magicalLineageState.enabled;
    magicalLineageLabel.textContent = ` ${magicalLineageSource.label} `;
    magicalLineageLabel.insertBefore(magicalLineageInput, magicalLineageLabel.firstChild);
  }

  let extendedScryingLabel = null;
  let extendedScryingInput = null;
  if (extendedScryingSource) {
    extendedScryingLabel = document.createElement("label");
    extendedScryingLabel.classList.add("checkbox");
    extendedScryingLabel.classList.add("nas-extended-scrying");
    extendedScryingInput = document.createElement("input");
    extendedScryingInput.type = "checkbox";
    extendedScryingInput.checked = extendedScryingState.enabled;
    extendedScryingLabel.textContent = ` ${extendedScryingSource.label} `;
    extendedScryingLabel.insertBefore(extendedScryingInput, extendedScryingLabel.firstChild);
  }

  if (maleficiumSource) {
    maleficiumLabel = document.createElement("label");
    maleficiumLabel.classList.add("checkbox");
    maleficiumLabel.classList.add("nas-maleficium");
    maleficiumInput = document.createElement("input");
    maleficiumInput.type = "checkbox";
    maleficiumInput.checked = maleficiumState.enabled;
    maleficiumLabel.textContent = ` ${maleficiumSource.label} `;
    maleficiumLabel.insertBefore(maleficiumInput, maleficiumLabel.firstChild);
  }

  const syncInput = () => {
    const payload = {
      ...priorFeatPayload,
      ...readFeatOptionsPayload(dialog, form),
    };
    if (source) {
      payload[source.id] = {
        enabled: state.enabled === true,
        mode: state.mode,
        selectedSpellUuid: state.selectedSpellUuid,
        selectedSpellLabel: state.selectedSpellLabel
      };
    }
    if (extendedScryingSource) {
      payload[extendedScryingSource.id] = {
        enabled: extendedScryingState.enabled === true
      };
    }
    if (magicalLineageSource) {
      payload[magicalLineageSource.id] = {
        enabled: magicalLineageState.enabled === true,
        selectedSpellUuid: magicalLineageState.selectedSpellUuid,
        selectedSpellLabel: magicalLineageState.selectedSpellLabel
      };
    }
    if (maleficiumSource) {
      payload[maleficiumSource.id] = {
        enabled: maleficiumState.enabled === true,
        damnationCount: Number.isFinite(maleficiumState.damnationCount) ? maleficiumState.damnationCount : 0
      };
    }
    DialogStateTracker.set(dialog.appId, FEAT_OPTIONS_KEY, payload);
    optionsInput.value = JSON.stringify(payload);
  };
  const updatePreview = () => {
    const infoContainer = form.querySelector(".nas-metamagic-info");
    if (!infoContainer) return;
    const metamagicListContainer = form.querySelector(".metamagic-options");
    const metamagicToggleInput = form.querySelector(`input[name="${METAMAGIC_FORM_KEY}"]`);
    updateMetamagicInfoSection(infoContainer, {
      baseSpellLevel: getSpellBaseLevel(dialog),
      baseSaveDc: getDialogSaveDc(dialog),
      metamagicEnabled: metamagicToggleInput?.checked !== false,
      options: getCurrentMetamagicOptions(dialog, form),
      featOptions: parseFeatOptionsInput(optionsInput.value),
      actionItem: dialog?.action?.item ?? null,
      selectedMetaNames: metamagicListContainer ? getSelectedMetaNames(metamagicListContainer) : [],
      dialog
    });
  };
  const syncUi = () => {
    if (source && modeButton) {
      const showDetails = state.enabled === true;
      modeButton.style.display = showDetails ? "" : "none";
      const isComplete = state.mode === modes.complete;
      modeButton.textContent = isComplete ? "C" : "I";
      modeButton.title = isComplete
        ? "Complete: +1 DC and metamagic cost reduction"
        : "Incomplete: +1 caster level";
    }
  };
  const persistState = () => {
    if (source) {
      void setEldritchResearcherState(actor, {
        mode: state.mode,
        selectedSpellUuid: state.selectedSpellUuid,
        selectedSpellLabel: state.selectedSpellLabel,
        selectedSpellName: state.selectedSpellLabel
      });
      setPersistentFeatureState(actor, source.id, state.enabled === true);
    }
    if (extendedScryingSource) {
      void setExtendedScryingState(actor, extendedScryingState.enabled);
    }
    if (magicalLineageSource) {
      void setMagicalLineageState(actor, {
        selectedSpellUuid: magicalLineageState.selectedSpellUuid,
        selectedSpellLabel: magicalLineageState.selectedSpellLabel,
        selectedSpellName: magicalLineageState.selectedSpellLabel
      });
      setPersistentFeatureState(actor, magicalLineageSource.id, magicalLineageState.enabled === true);
    }
    if (maleficiumSource) {
      void setMaleficiumState(actor, maleficiumState.enabled);
    }
  };
  const chooseSpell = async () => {
    const selection = await promptEldritchResearcherSpellSelection(actor, {
      currentSpellUuid: state.selectedSpellUuid,
      currentMode: state.mode,
      includeModeToggle: true,
      title: source.label,
      description: `Choose a spell for ${source.label}:`
    });
    if (!selection) return false;
    state.selectedSpellUuid = selection.selectedSpellUuid ?? "";
    state.selectedSpellLabel = selection.selectedSpellLabel ?? "";
    state.mode = selection.selectedMode === modes.complete ? modes.complete : modes.incomplete;
    return true;
  };

  if (source && input && modeButton) {
    input.addEventListener("change", async () => {
      state.enabled = input.checked === true;
      if (state.enabled && !state.selectedSpellUuid) {
        const didSelect = await chooseSpell();
        if (!didSelect && !state.selectedSpellUuid) {
          state.enabled = false;
          input.checked = false;
        }
      }
      syncUi();
      syncInput();
      updatePreview();
      persistState();
      requestAttackDialogAutoLayout(dialog);
    });
    modeButton.addEventListener("click", () => {
      state.mode = state.mode === modes.complete ? modes.incomplete : modes.complete;
      syncInput();
      syncUi();
      updatePreview();
      persistState();
    });
  }
  if (extendedScryingInput) {
    extendedScryingInput.addEventListener("change", () => {
      extendedScryingState.enabled = extendedScryingInput.checked === true;
      syncInput();
      updatePreview();
      persistState();
    });
  }
  if (magicalLineageSource && magicalLineageInput) {
    magicalLineageInput.addEventListener("change", async () => {
      magicalLineageState.enabled = magicalLineageInput.checked === true;
      if (magicalLineageState.enabled && !magicalLineageState.selectedSpellUuid) {
        const selection = await promptEldritchResearcherSpellSelection(actor, {
          currentSpellUuid: magicalLineageState.selectedSpellUuid
        });
        if (!selection && !magicalLineageState.selectedSpellUuid) {
          magicalLineageState.enabled = false;
          magicalLineageInput.checked = false;
        } else if (selection) {
          magicalLineageState.selectedSpellUuid = selection.selectedSpellUuid ?? "";
          magicalLineageState.selectedSpellLabel = selection.selectedSpellLabel ?? "";
        }
      }
      syncUi();
      syncInput();
      updatePreview();
      persistState();
      requestAttackDialogAutoLayout(dialog);
    });
  }
  if (maleficiumInput) {
    maleficiumInput.addEventListener("change", () => {
      maleficiumState.enabled = maleficiumInput.checked === true;
      syncInput();
      updatePreview();
      persistState();
    });
  }

  const actionItem = dialog?.action?.item ?? null;
  const eldritchConfigured = Boolean(state.selectedSpellUuid);
  const eldritchMatches = selectedSpellMatchesActionItem(actionItem, state.selectedSpellUuid);
  const showEldritch = !eldritchConfigured || eldritchMatches;
  const magicalLineageConfigured = Boolean(magicalLineageState.selectedSpellUuid);
  const magicalLineageMatches = selectedSpellMatchesActionItem(actionItem, magicalLineageState.selectedSpellUuid);
  const showMagicalLineage = !magicalLineageConfigured || magicalLineageMatches;

  syncUi();
  syncInput();
  updatePreview();
  if (source && label && showEldritch) {
    if (metamagicDropdown?.parentElement === container) {
      container.insertBefore(label, metamagicDropdown);
    } else {
      container.appendChild(label);
    }
  }
  if (extendedScryingLabel) {
    if (metamagicDropdown?.parentElement === container) {
      container.insertBefore(extendedScryingLabel, metamagicDropdown);
    } else {
      container.appendChild(extendedScryingLabel);
    }
  }
  if (magicalLineageLabel && showMagicalLineage) {
    const featureRow = featureContainer?.querySelector(".nas-metamagic-feature-row");
    if (featureRow) {
      featureContainer.style.display = "";
      featureRow.appendChild(magicalLineageLabel);
    } else if (featureContainer) {
      featureContainer.style.display = "";
      featureContainer.appendChild(magicalLineageLabel);
    } else if (metamagicDropdown?.parentElement === container) {
      container.insertBefore(magicalLineageLabel, metamagicDropdown);
    } else {
      container.appendChild(magicalLineageLabel);
    }
  }
  if (maleficiumLabel) {
    if (metamagicDropdown?.parentElement === container) {
      container.insertBefore(maleficiumLabel, metamagicDropdown);
    } else {
      container.appendChild(maleficiumLabel);
    }
  }
  requestAttackDialogAutoLayout(dialog);
}

async function renderSpellPerfectionStandaloneControl(dialog, form, actor, source) {
  if (!form || !actor || !source) return;
  form.querySelectorAll("label.nas-spell-perfection").forEach((node) => node.remove());

  const spellChoices = getEldritchResearcherSpellChoices(actor);
  if (!spellChoices.length) return;

  if (!ensureFeatOptionsInput(form, dialog)) return;

  const container = getFlagsContainer(form);
  const metamagicDropdown = container.querySelector("details.metamagic-dropdown");
  const payload = readFeatOptionsPayload(dialog, form);
  const storedState = await getSpellPerfectionState(actor);
  const persistentSpellPerfection = getDialogPersistentState(actor, source?.id, storedState.enabled === true);
  const payloadState = payload?.[source.id] && typeof payload[source.id] === "object" ? payload[source.id] : {};
  const state = {
    enabled: payloadState.enabled === true || (payloadState.enabled === undefined && persistentSpellPerfection),
    selectedSpellUuid: (payloadState.selectedSpellUuid ?? storedState.selectedSpellUuid ?? "").toString().trim(),
    selectedSpellLabel: (payloadState.selectedSpellLabel ?? storedState.selectedSpellLabel ?? "").toString().trim(),
    selectedSpellName: (payloadState.selectedSpellName ?? storedState.selectedSpellName ?? storedState.selectedSpellLabel ?? "").toString().trim()
  };
  if (!state.selectedSpellUuid && state.selectedSpellName) {
    const byName = spellChoices.find(
      (entry) => (entry?.label ?? "").toString().trim().toLowerCase() === state.selectedSpellName.trim().toLowerCase()
    ) ?? null;
    state.selectedSpellUuid = byName?.uuid || byName?.id || "";
    state.selectedSpellLabel = byName?.label ?? state.selectedSpellName;
  }
  if (!state.selectedSpellLabel && state.selectedSpellUuid) {
    state.selectedSpellLabel = resolveSpellLabelFromChoices(spellChoices, state.selectedSpellUuid, "");
  }
  const actionItem = dialog?.action?.item ?? null;
  const configured = Boolean(state.selectedSpellUuid);
  const matches = selectedSpellMatchesActionItem(actionItem, state.selectedSpellUuid);
  if (configured && !matches) return;

  const label = document.createElement("label");
  label.classList.add("checkbox");
  label.classList.add("nas-spell-perfection");
  const input = document.createElement("input");
  input.type = "checkbox";
  input.checked = state.enabled === true;
  label.textContent = ` ${source.label} `;
  label.insertBefore(input, label.firstChild);

  const refreshPreview = () => {
    const infoContainer = form.querySelector(".nas-metamagic-info");
    if (!infoContainer) return;
    const metamagicListContainer = form.querySelector(".metamagic-options");
    const metamagicToggleInput = form.querySelector(`input[name="${METAMAGIC_FORM_KEY}"]`);
    updateMetamagicInfoSection(infoContainer, {
      baseSpellLevel: getSpellBaseLevel(dialog),
      baseSaveDc: getDialogSaveDc(dialog),
      metamagicEnabled: metamagicToggleInput?.checked !== false,
      options: getCurrentMetamagicOptions(dialog, form),
      featOptions: getCurrentFeatOptions(dialog, form),
      actionItem: dialog?.action?.item ?? null,
      selectedMetaNames: metamagicListContainer ? getSelectedMetaNames(metamagicListContainer) : [],
      dialog
    });
  };

  const syncUi = () => {};

  const persist = () => {
    void setSpellPerfectionState(actor, {
      selectedSpellUuid: state.selectedSpellUuid,
      selectedSpellLabel: state.selectedSpellLabel,
      selectedSpellName: state.selectedSpellLabel
    });
    setPersistentFeatureState(actor, source.id, state.enabled === true);
  };

  const syncPayload = () => {
    const nextPayload = {
      ...readFeatOptionsPayload(dialog, form),
      [source.id]: {
        enabled: state.enabled === true,
        selectedSpellUuid: state.selectedSpellUuid,
        selectedSpellLabel: state.selectedSpellLabel,
        selectedSpellName: state.selectedSpellLabel
      }
    };
    writeFeatOptionsPayload(dialog, form, nextPayload);
  };

  const chooseSpell = async () => {
    const selection = await promptEldritchResearcherSpellSelection(actor, {
      currentSpellUuid: state.selectedSpellUuid
    });
    if (!selection) return false;
    state.selectedSpellUuid = selection.selectedSpellUuid ?? "";
    state.selectedSpellLabel = selection.selectedSpellLabel ?? "";
    return true;
  };

  input.addEventListener("change", async () => {
    state.enabled = input.checked === true;
    if (state.enabled && !state.selectedSpellUuid) {
      const didSelect = await chooseSpell();
      if (!didSelect && !state.selectedSpellUuid) {
        state.enabled = false;
        input.checked = false;
      }
    }
    syncUi();
    syncPayload();
    refreshPreview();
    persist();
    requestAttackDialogAutoLayout(dialog);
  });

  syncUi();
  syncPayload();
  refreshPreview();
  if (metamagicDropdown?.parentElement === container) {
    container.insertBefore(label, metamagicDropdown);
  } else {
    container.appendChild(label);
  }
  requestAttackDialogAutoLayout(dialog);
}

function getSpontaneousMetafocusFlagState(featItem) {
  const rawSelectedCount = getDictionaryNumber(featItem, SPONTANEOUS_METAFOCUS_SELECTED_FLAG, 1);
  const selectedCount = Number.isFinite(rawSelectedCount) && rawSelectedCount > 0
    ? Math.max(1, Math.trunc(rawSelectedCount))
    : 1;
  const selectedSpellNames = getDictionaryPrefixedStrings(featItem, SPONTANEOUS_METAFOCUS_SPELL_FLAG_PREFIX)
    .slice(0, selectedCount);
  return { selectedCount, selectedSpellNames };
}

function promptSpontaneousMetafocusSpellSelection(
  actor,
  {
    currentSpellNames = [],
    selectedCount = 1
  } = {}
) {
  const choices = getEldritchResearcherSpellChoices(actor);
  if (!choices.length) {
    ui.notifications.warn(localizeMetamagic("warnings.noNonSpelllikeSpellsFound"));
    return Promise.resolve(null);
  }

  const maxSelections = Math.max(1, Number(selectedCount) || 1);
  const selectedSet = new Set(
    (Array.isArray(currentSpellNames) ? currentSpellNames : [])
      .map((value) => (value ?? "").toString().trim())
      .map((value) => value.toLowerCase())
      .filter(Boolean)
  );
  const useCheckboxes = maxSelections > 1;
  const inputType = useCheckboxes ? "checkbox" : "radio";
  const inputName = "spontaneousMetafocusSpellChoice";
  const cards = choices
    .map((choice, index) => {
      const value = (choice.label ?? "").toString().trim();
      if (!value) return "";
      const checked = selectedSet.has(value.toLowerCase());
      const img = choice.img ? `<img src="${choice.img}" style="width:32px;height:32px;border-radius:4px;" />` : "";
      const safeLabel = foundry.utils.escapeHTML(choice.label ?? `Spell ${index + 1}`);
      return `
        <label class="nas-spontaneous-metafocus-spell-option" style="display:flex;align-items:center;gap:8px;padding:4px 0;">
          <input type="${inputType}" name="${inputName}" value="${value}" ${checked ? "checked" : ""} />
          ${img}
          <span>${safeLabel}</span>
        </label>
      `;
    })
    .join("");

  const title = localizeMetamagic("featureNames.spontaneousMetafocus");
  const description = useCheckboxes
    ? formatMetamagic("dialogs.chooseSpellsForName", { count: maxSelections, name: title })
    : formatMetamagic("dialogs.chooseSpellForName", { name: title });
  const content = `
    <form>
      <p style="margin:0 0 8px 0;">${foundry.utils.escapeHTML(description)}</p>
      <div style="max-height:320px;overflow-y:auto;border:1px solid #888;border-radius:4px;padding:6px 8px;">
        ${cards}
      </div>
    </form>
  `;

  return new Promise((resolve) => {
    new Dialog({
      title,
      content,
      buttons: {
        ok: {
          label: game.i18n.localize("NAS.common.buttons.select"),
          callback: (html) => {
            const values = html
              .find(`input[name="${inputName}"]:checked`)
              .map((_, el) => (el?.value ?? "").toString().trim())
              .get()
              .filter(Boolean);
            resolve(values);
          }
        },
        cancel: {
          label: game.i18n.localize("NAS.common.buttons.cancel"),
          callback: () => resolve(null)
        }
      },
      default: "ok",
      close: () => resolve(null),
      render: (html) => {
        if (!useCheckboxes) return;
        html.find(`input[name="${inputName}"]`).on("change", (event) => {
          const checked = html.find(`input[name="${inputName}"]:checked`).length;
          if (checked > maxSelections) {
            event.currentTarget.checked = false;
          }
        });
      }
    }).render(true);
  });
}

function resolveSpontaneousMetafocusNamesToSpells(spellChoices, spellNames) {
  const choices = Array.isArray(spellChoices) ? spellChoices : [];
  const byExactName = new Map();
  const byLowerName = new Map();
  const byUuid = new Map();
  const byId = new Map();
  for (const choice of choices) {
    const label = (choice?.label ?? "").toString().trim();
    const lower = label.toLowerCase();
    if (label && !byExactName.has(label)) byExactName.set(label, choice);
    if (lower && !byLowerName.has(lower)) byLowerName.set(lower, choice);
    const uuid = (choice?.uuid ?? "").toString().trim();
    const id = (choice?.id ?? "").toString().trim();
    if (uuid) byUuid.set(uuid, choice);
    if (id) byId.set(id, choice);
  }

  const resolved = [];
  const unresolvedNames = [];
  for (const rawName of Array.isArray(spellNames) ? spellNames : []) {
    const name = (rawName ?? "").toString().trim();
    if (!name) continue;
    const lower = name.toLowerCase();
    const selectedDocId = name.split(".").pop();
    const match = byExactName.get(name)
      ?? byLowerName.get(lower)
      ?? byUuid.get(name)
      ?? byId.get(name)
      ?? (selectedDocId ? byId.get(selectedDocId) : null);
    if (!match) {
      unresolvedNames.push(name);
      continue;
    }
    if (!resolved.some((entry) => entry.uuid === match.uuid || entry.id === match.id)) {
      resolved.push(match);
    }
  }

  return {
    resolvedSpellUuids: resolved.map((choice) => (choice?.uuid || choice?.id || "").toString()).filter(Boolean),
    resolvedSpellLabels: resolved.map((choice) => (choice?.label ?? "").toString().trim()).filter(Boolean),
    unresolvedNames
  };
}

async function saveSpontaneousMetafocusSpellNames(featItem, spellNames, selectedCount) {
  if (!featItem) return;
  const requiredCount = Math.max(1, Number(selectedCount) || 1);
  const chosen = Array.from(
    new Set(
      (Array.isArray(spellNames) ? spellNames : [])
        .map((value) => (value ?? "").toString().trim())
        .filter(Boolean)
    )
  ).slice(0, requiredCount);

  await setDictionaryEntries(featItem, {
    [SPONTANEOUS_METAFOCUS_SELECTED_FLAG]: String(requiredCount)
  });
  await setDictionaryPrefixedStrings(featItem, SPONTANEOUS_METAFOCUS_SPELL_FLAG_PREFIX, chosen);
}

async function ensureSpontaneousMetafocusSelectionConfigured(featItem, actor, spellChoices = []) {
  if (!featItem || !actor) return null;
  const current = getSpontaneousMetafocusFlagState(featItem);
  const requiredCount = Math.max(1, Number(current?.selectedCount ?? 1) || 1);
  const currentNames = Array.isArray(current?.selectedSpellNames) ? current.selectedSpellNames : [];
  const currentResolved = resolveSpontaneousMetafocusNamesToSpells(spellChoices, currentNames);
  if (
    currentResolved.unresolvedNames.length === 0
    && currentResolved.resolvedSpellUuids.length >= requiredCount
  ) {
    const normalizedNames = currentResolved.resolvedSpellLabels.slice(0, requiredCount);
    if (normalizedNames.join("|") !== currentNames.slice(0, requiredCount).join("|")) {
      await saveSpontaneousMetafocusSpellNames(featItem, normalizedNames, requiredCount);
    }
    return {
      selectedSpellUuids: currentResolved.resolvedSpellUuids.slice(0, requiredCount),
      selectedSpellLabels: currentResolved.resolvedSpellLabels.slice(0, requiredCount)
    };
  }

  const selected = await promptSpontaneousMetafocusSpellSelection(actor, {
    currentSpellNames: currentNames,
    selectedCount: requiredCount
  });
  if (!selected) return null;
  const chosenNames = Array.from(
    new Set(
      selected
        .map((value) => (value ?? "").toString().trim())
        .filter(Boolean)
    )
  ).slice(0, requiredCount);
  if (chosenNames.length !== requiredCount) return null;

  const resolved = resolveSpontaneousMetafocusNamesToSpells(spellChoices, chosenNames);
  if (resolved.resolvedSpellUuids.length !== requiredCount) return null;

  await saveSpontaneousMetafocusSpellNames(featItem, resolved.resolvedSpellLabels, requiredCount);
  return {
    selectedSpellUuids: resolved.resolvedSpellUuids.slice(0, requiredCount),
    selectedSpellLabels: resolved.resolvedSpellLabels.slice(0, requiredCount)
  };
}

async function renderSpontaneousMetafocusStandaloneControl(dialog, form, actor, source) {
  if (!form || !actor || !source) return;
  form.querySelectorAll("label.nas-spontaneous-metafocus").forEach((node) => node.remove());

  const spellChoices = getEldritchResearcherSpellChoices(actor);
  if (!spellChoices.length) return;

  if (!ensureFeatOptionsInput(form, dialog)) return;

  const container = getFlagsContainer(form);
  const metamagicDropdown = container.querySelector("details.metamagic-dropdown");
  const featureContainer = form.querySelector(".nas-metamagic-features");
  const payload = readFeatOptionsPayload(dialog, form);
  const payloadState = payload?.[source.id] && typeof payload[source.id] === "object" ? payload[source.id] : {};
  const featItem = source.itemUuid ? await fromUuid(source.itemUuid) : null;
  const flagState = getSpontaneousMetafocusFlagState(featItem);
  const resolvedFlagSelection = resolveSpontaneousMetafocusNamesToSpells(
    spellChoices,
    Array.isArray(flagState?.selectedSpellNames) ? flagState.selectedSpellNames : []
  );
  const selectedSpellUuids = Array.isArray(payloadState.selectedSpellUuids)
    ? payloadState.selectedSpellUuids.map((value) => (value ?? "").toString().trim()).filter(Boolean)
    : [];
  const baseSelectedSpellUuids = selectedSpellUuids.length
    ? selectedSpellUuids
    : Array.isArray(resolvedFlagSelection?.resolvedSpellUuids)
      ? resolvedFlagSelection.resolvedSpellUuids
      : [];
  const state = {
    enabled: payloadState.enabled === true,
    selectedSpellUuids: Array.from(new Set(baseSelectedSpellUuids)),
    selectedSpellLabels: Array.isArray(payloadState.selectedSpellLabels)
      ? payloadState.selectedSpellLabels.map((value) => (value ?? "").toString().trim()).filter(Boolean)
      : Array.isArray(resolvedFlagSelection?.resolvedSpellLabels)
        ? resolvedFlagSelection.resolvedSpellLabels
        : []
  };
  if (state.enabled && !state.selectedSpellUuids.length) {
    state.enabled = false;
  }
  const actionItem = dialog?.action?.item ?? null;
  const configuredSelectionPresent = state.selectedSpellUuids.length > 0;
  const selectedSpellMatchesAction = configuredSelectionPresent
    && state.selectedSpellUuids.some((uuid) => selectedSpellMatchesActionItem(actionItem, uuid));
  if (configuredSelectionPresent && !selectedSpellMatchesAction) return;

  const label = document.createElement("label");
  label.classList.add("checkbox");
  label.classList.add("nas-spontaneous-metafocus");
  const input = document.createElement("input");
  input.type = "checkbox";
  input.checked = state.enabled === true;
  label.textContent = ` ${source.label} `;
  label.insertBefore(input, label.firstChild);

  const refreshPreview = () => {
    const infoContainer = form.querySelector(".nas-metamagic-info");
    if (!infoContainer) return;
    const metamagicListContainer = form.querySelector(".metamagic-options");
    const metamagicToggleInput = form.querySelector(`input[name="${METAMAGIC_FORM_KEY}"]`);
    updateMetamagicInfoSection(infoContainer, {
      baseSpellLevel: getSpellBaseLevel(dialog),
      baseSaveDc: getDialogSaveDc(dialog),
      metamagicEnabled: metamagicToggleInput?.checked !== false,
      options: getCurrentMetamagicOptions(dialog, form),
      featOptions: getCurrentFeatOptions(dialog, form),
      actionItem: dialog?.action?.item ?? null,
      selectedMetaNames: metamagicListContainer ? getSelectedMetaNames(metamagicListContainer) : [],
      dialog
    });
  };

  const syncPayload = () => {
    const nextPayload = {
      ...readFeatOptionsPayload(dialog, form),
      [source.id]: {
        enabled: state.enabled === true,
        selectedSpellUuids: [...state.selectedSpellUuids],
        selectedSpellLabels: [...state.selectedSpellLabels]
      }
    };
    writeFeatOptionsPayload(dialog, form, nextPayload);
  };

  const syncUi = () => {};

  const chooseSpells = async () => {
    if (!featItem || featItem.actor?.id !== actor.id) {
      ui.notifications.warn(localizeMetamagic("warnings.spontaneousMetafocusFeatNotFound"));
      return false;
    }
    const selected = await ensureSpontaneousMetafocusSelectionConfigured(featItem, actor, spellChoices);
    if (!selected?.selectedSpellUuids?.length) return false;
    state.selectedSpellUuids = [...selected.selectedSpellUuids];
    state.selectedSpellLabels = [...(selected.selectedSpellLabels ?? [])];
    return true;
  };

  input.addEventListener("change", async () => {
    state.enabled = input.checked === true;
    if (state.enabled) {
      const didChoose = await chooseSpells();
      if (!didChoose && !state.selectedSpellUuids.length) {
        state.enabled = false;
        input.checked = false;
      }
    }
    syncUi();
    syncPayload();
    refreshPreview();
    requestAttackDialogAutoLayout(dialog);
  });

  syncUi();
  syncPayload();
  refreshPreview();
  const featureRow = featureContainer?.querySelector(".nas-metamagic-feature-row");
  if (featureRow) {
    featureContainer.style.display = "";
    featureRow.appendChild(label);
  } else if (featureContainer) {
    featureContainer.style.display = "";
    featureContainer.appendChild(label);
  } else if (metamagicDropdown?.parentElement === container) {
    container.insertBefore(label, metamagicDropdown);
  } else {
    container.appendChild(label);
  }
  requestAttackDialogAutoLayout(dialog);
}

function formatLimitedUsesSuffix(source) {
  if (!source?.limited) return "";
  const usesValue = Number(source?.usesValue);
  const usesMax = Number(source?.usesMax);
  if (Number.isFinite(usesValue) && Number.isFinite(usesMax) && usesMax >= 0) {
    const safeValue = Math.max(0, Math.trunc(usesValue));
    const safeMax = Math.max(0, Math.trunc(usesMax));
    return ` ${safeValue}/${safeMax}`;
  }
  return " (limited uses)";
}

function renderTraitModifierControls(dialog, form, traitSources = []) {
  if (!form) return;
  const actor = getActionActor(dialog);
  const sources = Array.isArray(traitSources) ? traitSources : [];
  const existing = form.querySelector(".nas-trait-modifiers");
  if (existing) existing.remove();
  form.querySelectorAll("label.nas-trait-modifier").forEach((element) => element.remove());
  const existingInput = form.querySelector(`input[name="${TRAIT_OPTIONS_KEY}"]`);
  if (existingInput) existingInput.remove();
  if (!sources.length) return;

  const container = getFlagsContainer(form);
  const metamagicDropdown = container.querySelector("details.metamagic-dropdown");
  const featureContainer = form.querySelector(".nas-metamagic-features");
  const featureRow = featureContainer?.querySelector(".nas-metamagic-feature-row");
  const metamagicToggleInput = form.querySelector(`input[name="${METAMAGIC_FORM_KEY}"]`);
  const metamagicListContainer = form.querySelector(".metamagic-options");
  const optionsInput = document.createElement("input");
  optionsInput.type = "hidden";
  optionsInput.name = TRAIT_OPTIONS_KEY;
  if (metamagicDropdown?.parentElement === container) {
    container.insertBefore(optionsInput, metamagicDropdown);
  } else {
    container.appendChild(optionsInput);
  }

  const storedSelections = parseTraitSelectionInput(DialogStateTracker.get(dialog.appId, TRAIT_OPTIONS_KEY));
  const selected = new Set(storedSelections);
  const nodesById = new Map();
  for (const source of sources) {
    if (
      source?.persistent === true
      && getDialogPersistentState(actor, source?.id, source?.enabled === true) === true
      && source?.requiresSpellSelection !== true
    ) {
      selected.add(source.id);
    }
  }

  const hasMetamagicIntent = () => {
    if (metamagicToggleInput?.checked !== true) return false;
    if (!metamagicListContainer) return false;
    return getSelectedMetaNames(metamagicListContainer).length > 0;
  };
  const shouldShowSource = (source) => {
    if (source?.requiresMetamagicIntent !== true) return true;
    return hasMetamagicIntent();
  };

  const syncInput = () => {
    const next = JSON.stringify(Array.from(selected));
    DialogStateTracker.set(dialog.appId, TRAIT_OPTIONS_KEY, next);
    optionsInput.value = next;
  };
  const resolveTraitItemFromSource = (source) => {
    if (!actor?.items) return null;
    const sourceId = (source?.itemId ?? "").toString().trim();
    if (sourceId) {
      const byId = actor.items.get?.(sourceId);
      if (byId) return byId;
    }
    const sourceUuid = (source?.itemUuid ?? "").toString().trim();
    const sourceUuidId = sourceUuid ? sourceUuid.split(".").pop() : "";
    if (sourceUuidId) {
      const byUuidId = actor.items.get?.(sourceUuidId);
      if (byUuidId) return byUuidId;
    }
    return null;
  };
  let traitMetaSyncLock = false;
  const syncTransmuterExtendMutualExclusion = ({ fromTraitChange = false } = {}) => {
    if (traitMetaSyncLock) return;
    const transmuterNode = nodesById.get(TRANSMUTER_OF_KORADA_ID);
    if (!transmuterNode || !metamagicListContainer) return;
    const extendCheckbox = metamagicListContainer.querySelector(
      `input[type="checkbox"][data-meta-name="${ExtendSpellDef.name}"]`
    );
    if (!extendCheckbox) return;
    const transmuterSelected = selected.has(TRANSMUTER_OF_KORADA_ID) && transmuterNode.input.checked === true;
    const extendSelected = extendCheckbox.checked === true;
    if (!transmuterSelected || !extendSelected) return;
    traitMetaSyncLock = true;
    if (fromTraitChange) {
      extendCheckbox.checked = false;
      extendCheckbox.dispatchEvent(new Event("change", { bubbles: true }));
    } else {
      selected.delete(TRANSMUTER_OF_KORADA_ID);
      transmuterNode.input.checked = false;
      syncInput();
    }
    traitMetaSyncLock = false;
  };

  sources
    .slice()
    .sort((a, b) => (a?.label ?? "").localeCompare(b?.label ?? ""))
    .forEach((source) => {
      if (source?.id === TRANSMUTER_OF_KORADA_ID && source?.hasRemaining !== true) {
        return;
      }
      const label = document.createElement("label");
      label.classList.add("checkbox");
      label.classList.add("nas-trait-modifier");
      const input = document.createElement("input");
      input.type = "checkbox";
      input.dataset.traitId = source.id;
      input.checked = selected.has(source.id);

      const limitedSuffix = formatLimitedUsesSuffix(source);
      label.textContent = ` ${source.label}${limitedSuffix} `;
      label.insertBefore(input, label.firstChild);
      if (source?.id === WAYANG_SPELLHUNTER_ID && featureRow) {
        featureContainer.style.display = "";
        featureRow.appendChild(label);
      } else if (metamagicDropdown?.parentElement === container) {
        container.insertBefore(label, metamagicDropdown);
      } else {
        container.appendChild(label);
      }

      input.addEventListener("change", async () => {
        if (input.checked) {
          if (source.id === TRANSMUTER_OF_KORADA_ID && source?.requiresSpellSelection === true) {
            const traitItem = resolveTraitItemFromSource(source);
            const selection = await promptTransmuterOfKoradaSpellSelection(traitItem);
            if (!selection?.selectedSpellName) {
              input.checked = false;
              selected.delete(source.id);
              syncInput();
              requestAttackDialogAutoLayout(dialog);
              return;
            }
            if (traitItem) {
              await setTransmuterOfKoradaSelectedSpellName(traitItem, selection.selectedSpellName);
            }
            source.selectedSpellName = selection.selectedSpellName;
            source.requiresSpellSelection = false;
          }
          if (source.id === WAYANG_SPELLHUNTER_ID && source?.requiresSpellSelection === true) {
            const traitItem = resolveTraitItemFromSource(source);
            const selection = await promptWayangSpellhunterSpellSelection(actor, {
              currentSpellName: source?.selectedSpellName ?? "",
              title: source?.label ?? "Wayang Spellhunter (Minata)"
            });
            if (!selection?.selectedSpellName) {
              input.checked = false;
              selected.delete(source.id);
              syncInput();
              requestAttackDialogAutoLayout(dialog);
              return;
            }
            if (traitItem) {
              await setWayangSpellhunterSelectedSpellName(traitItem, selection.selectedSpellName);
            }
            source.selectedSpellName = selection.selectedSpellName;
            source.selectedSpellUuid = selection.selectedSpellUuid ?? "";
            source.selectedSpellMatches = selectedSpellMatchesActionItem(
              dialog?.action?.item ?? null,
              source.selectedSpellUuid
            );
            source.requiresSpellSelection = !source.selectedSpellUuid;
          }
          selected.add(source.id);
        } else {
          selected.delete(source.id);
        }
        if (source.id === WAYANG_SPELLHUNTER_ID) {
          setPersistentFeatureState(actor, source.id, input.checked === true);
        }
        syncInput();
        if (source.id === TRANSMUTER_OF_KORADA_ID && input.checked === true) {
          syncTransmuterExtendMutualExclusion({ fromTraitChange: true });
        }
      });

      nodesById.set(source.id, { source, label, input });
    });

  const syncVisibility = () => {
    let changed = false;
    for (const [id, node] of nodesById.entries()) {
      const visible = shouldShowSource(node.source);
      node.label.style.display = visible ? "" : "none";
      node.input.disabled = !visible;
      if (!visible && selected.has(id)) {
        if (node.source?.preserveSelectionWhenHidden === true) {
          continue;
        }
        selected.delete(id);
        node.input.checked = false;
        changed = true;
      }
    }
    if (changed) {
      syncInput();
    } else if (!optionsInput.value) {
      syncInput();
    }
    requestAttackDialogAutoLayout(dialog);
  };

  if (metamagicToggleInput) {
    metamagicToggleInput.addEventListener("change", syncVisibility);
  }
  if (metamagicListContainer) {
    metamagicListContainer.addEventListener("change", () => {
      syncVisibility();
      syncTransmuterExtendMutualExclusion({ fromTraitChange: false });
    });
  }

  syncVisibility();
  syncTransmuterExtendMutualExclusion({ fromTraitChange: false });

  requestAttackDialogAutoLayout(dialog);
}

function getSlotIncreaseForSelection(metaName, options, baseSpellLevel, featOptions = {}) {
  const normalized = (metaName ?? "").toString();
  if (!normalized) return 0;
  if (normalized === "Extend Spell" && featOptions?.maskFocus?.enabled === true) {
    return 0;
  }
  if (normalized === "Reach Spell") {
    const steps = Number(options?.reachSpellSteps ?? 1);
    return Number.isFinite(steps) && steps > 0 ? steps : 1;
  }
  if (normalized === "Heighten Spell") {
    const target = Number(options?.heightenSpellLevel ?? 0);
    const base = Number(baseSpellLevel ?? 0);
    if (!Number.isFinite(target) || !Number.isFinite(base)) return 0;
    return Math.max(0, target - base);
  }
  return Number(METAMAGIC_SLOT_COSTS[normalized] ?? 0);
}

function selectedSpellMatchesActionItem(actionItem, selectedSpellUuid) {
  if (!actionItem || !selectedSpellUuid) return false;
  const actionUuid = (actionItem?.uuid ?? "").toString().trim();
  const actionId = (actionItem?.id ?? "").toString().trim();
  if (actionUuid && actionUuid === selectedSpellUuid) return true;
  const selectedId = selectedSpellUuid.split(".").pop();
  if (selectedId && actionId && actionId === selectedId) return true;
  return false;
}

function isEvilDescriptorSpellItem(actionItem) {
  return spellItemHasEvilDescriptor(actionItem);
}

function getPreviewMetamagicSlotAdjustment({
  selectedMetaNames = [],
  options = {},
  featOptions = {},
  actionItem = null,
  hasSlotSurchargeMetamagic = false
} = {}) {
  if (!Array.isArray(selectedMetaNames) || selectedMetaNames.length === 0) return 0;
  if (!hasSlotSurchargeMetamagic) return 0;
  if (!featOptions || typeof featOptions !== "object") return 0;

  let adjustment = 0;
  const eldritchResearcher = featOptions.eldritchResearcher;
  if (
    eldritchResearcher?.enabled === true
    && eldritchResearcher?.mode === "complete"
    && selectedSpellMatchesActionItem(actionItem, eldritchResearcher?.selectedSpellUuid ?? "")
  ) {
    adjustment -= 1;
  }

  const magicalLineage = featOptions.magicalLineage;
  if (
    magicalLineage?.enabled === true
    && selectedSpellMatchesActionItem(actionItem, magicalLineage?.selectedSpellUuid ?? "")
  ) {
    adjustment -= 1;
  }

  const maleficium = featOptions.maleficium;
  const damnationCount = Number(maleficium?.damnationCount ?? 0);
  if (
    maleficium?.enabled === true
    && Number.isFinite(damnationCount)
    && damnationCount >= 2
    && isEvilDescriptorSpellItem(actionItem)
  ) {
    adjustment -= 1;
  }
  const features = getFeatureStatesFromOptions(options);
  if (features?.retribution === true) {
    adjustment -= 1;
  }
  return adjustment;
}

function getMaskFocusDisplayNameFromDialog(dialog) {
  const rawEl = dialog?.element;
  const root = elementFromHtmlLike(rawEl);
  if (!root?.querySelector) return "";
  const label = root.querySelector("form label.nas-mask-focus[data-nas-mf-label]");
  return (label?.getAttribute?.("data-nas-mf-label") ?? "").toString().trim();
}

function updateMetamagicInfoSection(infoContainer, {
  selectedMetaNames = [],
  options = {},
  featOptions = {},
  actionItem = null,
  baseSpellLevel = 0,
  baseSaveDc = null,
  metamagicEnabled = true,
  dialog = null
} = {}) {
  if (!infoContainer) return;
  const previewModeRaw = game?.settings?.get?.(MODULE.ID, "metamagicPreviewMode");
  const previewMode = [
    METAMAGIC_PREVIEW_MODE_NONE,
    METAMAGIC_PREVIEW_MODE_CONCISE,
    METAMAGIC_PREVIEW_MODE_DETAILED
  ].includes(previewModeRaw)
    ? previewModeRaw
    : METAMAGIC_PREVIEW_MODE_CONCISE;
  if (previewMode === METAMAGIC_PREVIEW_MODE_NONE) {
    infoContainer.innerHTML = "";
    infoContainer.style.display = "none";
    if (dialog) requestAttackDialogAutoLayout(dialog);
    return;
  }
  infoContainer.style.display = "";

  const selected = metamagicEnabled ? selectedMetaNames : [];
  const heightenIncrease = selected.includes("Heighten Spell")
    ? getSlotIncreaseForSelection("Heighten Spell", options, baseSpellLevel, featOptions)
    : 0;
  const otherIncrease = selected.reduce(
    (total, name) => total + (name === "Heighten Spell" ? 0 : getSlotIncreaseForSelection(name, options, baseSpellLevel, featOptions)),
    0
  );
  const safeBase = Number.isFinite(Number(baseSpellLevel)) ? Number(baseSpellLevel) : 0;
  const features = getFeatureStatesFromOptions(options);
  const timelessSoulEnabled = features.timelessSoul === true;
  const hasQuickenSelected = selected.includes("Quicken Spell");
  const timelessSoulApplies = timelessSoulEnabled && hasQuickenSelected;
  const quickenRawIncrease = hasQuickenSelected ? getSlotIncreaseForSelection("Quicken Spell", options, safeBase, featOptions) : 0;
  const timelessSoulReduction = timelessSoulApplies ? Math.min(1, Math.max(0, quickenRawIncrease)) : 0;
  const timelessSoulFixedQuicken = Math.max(0, quickenRawIncrease - timelessSoulReduction);
  const otherExcludingQuicken = timelessSoulApplies
    ? Math.max(0, otherIncrease - quickenRawIncrease)
    : otherIncrease;
  const hasSlotSurchargeMetamagic = otherExcludingQuicken > 0;
  const slotAdjustment = getPreviewMetamagicSlotAdjustment({
    selectedMetaNames: selected,
    options,
    featOptions,
    actionItem,
    hasSlotSurchargeMetamagic
  });
  const reducedOther = timelessSoulApplies
    ? timelessSoulFixedQuicken + Math.max(0, otherExcludingQuicken + slotAdjustment)
    : Math.max(0, otherIncrease + slotAdjustment);
  let consumedIncrease = heightenIncrease + reducedOther;
  let spellPerfectionWaiver = 0;
  const spellPerfection = featOptions?.spellPerfection;
  if (
    spellPerfection?.enabled === true
    && selectedSpellMatchesActionItem(actionItem, spellPerfection?.selectedSpellUuid ?? "")
  ) {
    spellPerfectionWaiver = selected.reduce((max, metaName) => {
      if ((metaName ?? "").toString() === "Quicken Spell" && timelessSoulApplies) return max;
      const increase = Number(getSlotIncreaseForSelection(metaName, options, safeBase, featOptions));
      if (!Number.isFinite(increase) || increase <= 0) return max;
      return Math.max(max, increase);
    }, 0);
    consumedIncrease = Math.max(0, consumedIncrease - spellPerfectionWaiver);
  }
  const maleficiumCount = Number(featOptions?.maleficium?.damnationCount ?? 0);
  const hasMaleficiumFloor =
    featOptions?.maleficium?.enabled === true
    && Number.isFinite(maleficiumCount)
    && maleficiumCount >= 2
    && isEvilDescriptorSpellItem(actionItem)
    && selected.length > 0;
  if (hasMaleficiumFloor) {
    consumedIncrease = Math.max(consumedIncrease, 1);
  }
  const metamagicMasteryOn = features.metamagicMastery === true;
  const resultLevel = metamagicMasteryOn ? safeBase : safeBase + consumedIncrease;
  const activeFeatures = Object.entries(features)
    .filter(([, value]) => value === true)
    .map(([key]) => key.replace(/([A-Z])/g, " $1").replace(/^./, (m) => m.toUpperCase()));

  const detailLines = [];
  if (selected.includes("Reach Spell")) {
    const steps = Number(options?.reachSpellSteps ?? 1);
    detailLines.push(formatMetamagic("preview.reachDetail", {
      steps: Number.isFinite(steps) ? steps : 1
    }));
  }
  if (selected.includes("Heighten Spell")) {
    const level = Number(options?.heightenSpellLevel ?? 0);
    if (Number.isFinite(level) && level > 0) {
      detailLines.push(formatMetamagic("preview.heightenDetail", { level }));
    }
  }
  if (featOptions?.maskFocus?.enabled === true && selected.includes(ExtendSpellDef.name)) {
    const maskFocusName = getMaskFocusDisplayNameFromDialog(dialog);
    if (maskFocusName) detailLines.push(maskFocusName);
  }

  const selectedText = selected.length ? selected.join(", ") : game.i18n.localize("NAS.common.labels.none");
  const detailsText = detailLines.length ? detailLines.join(" | ") : game.i18n.localize("NAS.common.labels.none");
  const activeFeaturesText = activeFeatures.length ? activeFeatures.join(", ") : game.i18n.localize("NAS.common.labels.none");
  const effectiveSpellLevel = safeBase + heightenIncrease;
  const hasSlotMathModifiers = otherIncrease !== 0 || slotAdjustment !== 0 || timelessSoulReduction > 0;
  const reducedOtherText = timelessSoulApplies
    ? `${timelessSoulFixedQuicken + Math.max(0, otherExcludingQuicken + slotAdjustment)}`
    : `${otherIncrease + slotAdjustment < 0 ? 0 : otherIncrease + slotAdjustment}`;
  const slotModsText = hasSlotMathModifiers
    ? timelessSoulApplies
      ? formatMetamagic("preview.slotModsTimeless", {
        otherIncrease,
        quickenFixed: timelessSoulFixedQuicken,
        slotAdjustment: `${slotAdjustment >= 0 ? "+" : ""}${slotAdjustment}`,
        reduced: reducedOtherText
      })
      : formatMetamagic("preview.slotModsDefault", {
        otherIncrease,
        slotAdjustment: `${slotAdjustment >= 0 ? "+" : ""}${slotAdjustment}`,
        reduced: reducedOtherText
      })
    : "";
  const spellPerfectionText = spellPerfectionWaiver > 0
    ? formatMetamagic("preview.spellPerfectionWaiver", { value: spellPerfectionWaiver })
    : "";
  const arcaneBloodlineEnabled = features.arcaneBloodline === true;
  const hasHeighten = selected.includes("Heighten Spell");
  const canApplyArcaneDc =
    arcaneBloodlineEnabled && !hasHeighten && otherIncrease > 0 && !metamagicMasteryOn;
  let saveDcText = localizeMetamagic("preview.saveDcNoChange");
  if (canApplyArcaneDc) {
    if (Number.isFinite(baseSaveDc)) {
      saveDcText = formatMetamagic("preview.saveDcArcaneCalc", { base: baseSaveDc, total: baseSaveDc + 1 });
    } else {
      saveDcText = localizeMetamagic("preview.saveDcArcaneBonus");
    }
  }
  const mmPreviewLine =
    metamagicMasteryOn && selected.length
      ? `<div>${game.i18n.format("NAS.metamagic.metamagicMastery.previewDebit", {
          count: Math.max(1, consumedIncrease)
        })}</div>`
      : "";

  if (previewMode === METAMAGIC_PREVIEW_MODE_CONCISE) {
    const conciseSaveDcText = saveDcText.replace(/\s*\(([^)]+)\)\s*/g, " - $1");
    infoContainer.innerHTML = `
      <div><strong>${localizeMetamagic("preview.title")}</strong></div>
      <div>${formatMetamagic("preview.selected", { value: selectedText })}</div>
      <div>${formatMetamagic("preview.options", { value: detailsText })}</div>
      <div>${formatMetamagic("preview.requiredSlotLevel", { value: resultLevel })}</div>
      ${mmPreviewLine}
      <div>${conciseSaveDcText}</div>
      <div>${formatMetamagic("preview.activeModifiers", { value: activeFeaturesText })}</div>
    `;
    if (dialog) requestAttackDialogAutoLayout(dialog);
    return;
  }

  infoContainer.innerHTML = `
    <div><strong>${localizeMetamagic("preview.title")}</strong></div>
    <div>${formatMetamagic("preview.selected", { value: selectedText })}</div>
    <div>${formatMetamagic("preview.options", { value: detailsText })}</div>
    <div>${formatMetamagic("preview.effectiveSpellLevel", { value: effectiveSpellLevel })}</div>
    <div>${formatMetamagic("preview.requiredSlotLevelUses", { value: resultLevel })}</div>
    ${mmPreviewLine}
    ${hasSlotMathModifiers ? `<div>${slotModsText}</div>` : ""}
      ${spellPerfectionText ? `<div>${spellPerfectionText}</div>` : ""}
    <div>${saveDcText}</div>
    <div>${formatMetamagic("preview.activeModifiers", { value: activeFeaturesText })}</div>
  `;
  if (dialog) requestAttackDialogAutoLayout(dialog);
}

function renderFeatureSection(
  dialog,
  form,
  featureContainer,
  optionsInput,
  actor,
  featureSources,
  listContainer,
  infoContainer,
  actionItem,
  baseSpellLevel,
  baseSaveDc,
  maskFocusSource = null,
  metamagicToggleInput = null
) {
  if (!featureContainer) return;
  const sources = Array.isArray(featureSources) ? featureSources : [];
  featureContainer.replaceChildren();
  if (!sources.length && !maskFocusSource) {
    featureContainer.style.display = 'none';
    return;
  }

  const title = document.createElement('div');
  title.innerHTML = `<strong>${localizeMetamagic("preview.modifiersTitle")}</strong>`;
  featureContainer.appendChild(title);

  const featureRow = document.createElement("div");
  featureRow.classList.add("nas-metamagic-feature-row");
  featureRow.style.display = "grid";
  featureRow.style.gridTemplateColumns = "repeat(3, minmax(0, 1fr))";
  featureRow.style.columnGap = "8px";
  featureRow.style.alignItems = "center";
  featureContainer.appendChild(featureRow);

  const storedOptions = DialogStateTracker.get(dialog.appId, METAMAGIC_OPTIONS_KEY) || {};
  let featureStates = getFeatureStatesFromOptions(storedOptions);
  const featureNodes = [];

  sources.forEach((source) => {
    const label = document.createElement('label');
    label.classList.add('checkbox');
    label.style.flex = "0 0 auto";
    label.dataset.featureId = source.id;
    const input = document.createElement('input');
    input.type = 'checkbox';
    input.dataset.featureId = source.id;
    const persistentValue = source.persistent ? getPersistentFeatureState(actor, source.id) : undefined;
    let initialChecked = false;
    if (featureStates[source.id] !== undefined) {
      initialChecked = featureStates[source.id] === true;
    } else if (persistentValue !== undefined) {
      initialChecked = persistentValue === true;
    } else {
      initialChecked = source.defaultEnabled === true;
    }
    input.checked = Boolean(initialChecked);

    input.addEventListener('change', () => {
      const currentOptions = DialogStateTracker.get(dialog.appId, METAMAGIC_OPTIONS_KEY) || {};
      const nextStates = {
        ...getFeatureStatesFromOptions(currentOptions),
        [source.id]: input.checked
      };
      const nextOptions = setFeatureStatesInOptions(currentOptions, nextStates);
      DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, nextOptions);
      updateMetamagicOptionsInput(optionsInput, nextOptions);
      updateMetamagicInfoSection(infoContainer, {
        baseSpellLevel,
        baseSaveDc,
        metamagicEnabled: DialogStateTracker.get(dialog.appId, METAMAGIC_FORM_KEY) !== false,
        options: nextOptions,
        featOptions: getCurrentFeatOptions(dialog, form),
        actionItem,
        selectedMetaNames: getSelectedMetaNames(listContainer),
        dialog
      });
    });

    const labelText = `${source.label}${formatLimitedUsesSuffix(source)}`;
    label.textContent = ` ${labelText} `;
    label.insertBefore(input, label.firstChild);
    featureRow.appendChild(label);
    featureNodes.push({ source, label, input, requiredWasVisible: false });
  });

  if (maskFocusSource && listContainer) {
    ensureFeatOptionsInput(form, dialog);
    wireMaskFocusExtendCoupling(form, dialog, listContainer);
    const maskFocusLabel = document.createElement("label");
    maskFocusLabel.classList.add("checkbox");
    maskFocusLabel.classList.add("nas-mask-focus");
    maskFocusLabel.style.flex = "0 0 auto";
    const maskFocusInput = document.createElement("input");
    maskFocusInput.type = "checkbox";
    const featOpts = readFeatOptionsPayload(dialog, form);
    maskFocusInput.checked = featOpts[MASK_FOCUS_ID]?.enabled === true;
    maskFocusInput.addEventListener("change", () => {
      const enabled = maskFocusInput.checked === true;
      patchFeatOptionsMaskFocus(dialog, form, enabled);
      if (enabled) {
        const mmToggle = form.querySelector(`input[name="${METAMAGIC_FORM_KEY}"]`);
        const dataInp = form.querySelector(`input[name="${METAMAGIC_NAMES_KEY}"]`);
        if (mmToggle && !mmToggle.checked) {
          mmToggle.checked = true;
          mmToggle.dispatchEvent(new Event("change", { bubbles: true }));
        }
        if (listContainer && dataInp) {
          const extendCb = listContainer.querySelector(`input[type="checkbox"][data-meta-name="${ExtendSpellDef.name}"]`);
          if (extendCb && !extendCb.checked) {
            extendCb.checked = true;
            extendCb.dispatchEvent(new Event("change", { bubbles: true }));
          } else {
            updateMetamagicNames(dataInp, mmToggle?.checked !== false, listContainer);
            if (infoContainer) {
              updateMetamagicInfoSection(infoContainer, {
                baseSpellLevel,
                baseSaveDc,
                metamagicEnabled: mmToggle?.checked !== false,
                options: getCurrentMetamagicOptions(dialog, form),
                featOptions: readFeatOptionsPayload(dialog, form),
                actionItem,
                selectedMetaNames: getSelectedMetaNames(listContainer),
                dialog
              });
            }
          }
        }
      } else if (infoContainer) {
        updateMetamagicInfoSection(infoContainer, {
          baseSpellLevel,
          baseSaveDc,
          metamagicEnabled: metamagicToggleInput?.checked !== false,
          options: getCurrentMetamagicOptions(dialog, form),
          featOptions: readFeatOptionsPayload(dialog, form),
          actionItem,
          selectedMetaNames: getSelectedMetaNames(listContainer),
          dialog
        });
      }
      requestAttackDialogAutoLayout(dialog);
    });
    const mfLabelText = `${maskFocusSource.label}${formatLimitedUsesSuffix(maskFocusSource)}`;
    maskFocusLabel.setAttribute("data-nas-mf-label", maskFocusSource.label ?? "");
    maskFocusLabel.textContent = ` ${mfLabelText} `;
    maskFocusLabel.insertBefore(maskFocusInput, maskFocusLabel.firstChild);
    featureRow.appendChild(maskFocusLabel);
  }

  featureStates = { ...featureStates };
  for (const source of sources) {
    const persistentValue = source.persistent ? getPersistentFeatureState(actor, source.id) : undefined;
    if (featureStates[source.id] === undefined && persistentValue !== undefined) {
      featureStates[source.id] = Boolean(persistentValue);
    } else if (featureStates[source.id] === undefined && source.defaultEnabled === true) {
      featureStates[source.id] = true;
    }
  }

  const updatedOptions = setFeatureStatesInOptions(storedOptions, featureStates);
  DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, updatedOptions);
  updateMetamagicOptionsInput(optionsInput, updatedOptions);

  const syncFeatureVisibility = () => {
    if (!featureNodes.length) return;
    const metamagicEnabled = metamagicToggleInput?.checked !== false;
    const selectedMetaNames = listContainer ? getSelectedMetaNames(listContainer) : [];
    const normalizedSelected = new Set(
      selectedMetaNames
        .map((name) => resolveMetamagicNameFromDatabase(name) ?? name)
        .map((name) => (name ?? "").toString().trim())
        .filter(Boolean)
    );

    let optionsChanged = false;
    let nextOptions = null;
    const currentOptions = DialogStateTracker.get(dialog.appId, METAMAGIC_OPTIONS_KEY) || {};
    const nextStates = { ...getFeatureStatesFromOptions(currentOptions) };
    for (const node of featureNodes) {
      const requiredNames = Array.isArray(node?.source?.requiredMetamagicNames)
        ? node.source.requiredMetamagicNames
        : [];
      if (!requiredNames.length) {
        node.label.style.display = "";
        node.input.disabled = false;
        continue;
      }
      const normalizedRequired = requiredNames
        .map((name) => resolveMetamagicNameFromDatabase(name) ?? name)
        .map((name) => (name ?? "").toString().trim())
        .filter(Boolean);
      const visible = metamagicEnabled && normalizedRequired.some((name) => normalizedSelected.has(name));
      if (
        node?.source?.autoCheckOnRequiredMetamagicToggle === true
        && visible
        && node.requiredWasVisible === false
        && node.input.checked !== true
      ) {
        node.input.checked = true;
        nextStates[node.source.id] = true;
        optionsChanged = true;
      }
      node.label.style.display = visible ? "" : "none";
      node.input.disabled = !visible;
      node.requiredWasVisible = visible;
    }
    if (optionsChanged) {
      nextOptions = setFeatureStatesInOptions(currentOptions, nextStates);
      DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, nextOptions);
      updateMetamagicOptionsInput(optionsInput, nextOptions);
      updateMetamagicInfoSection(infoContainer, {
        baseSpellLevel,
        baseSaveDc,
        metamagicEnabled,
        options: nextOptions,
        featOptions: getCurrentFeatOptions(dialog, form),
        actionItem,
        selectedMetaNames,
        dialog
      });
    }
    requestAttackDialogAutoLayout(dialog);
  };

  if (metamagicToggleInput) {
    metamagicToggleInput.addEventListener("change", syncFeatureVisibility);
  }
  if (listContainer) {
    listContainer.addEventListener("change", syncFeatureVisibility);
  }
  syncFeatureVisibility();
  featureContainer.style.display = '';
}

function renderHealersBlessingStandaloneControl(dialog, form, actor, sources) {
  if (!form || !actor) return;
  const source = Array.isArray(sources) ? sources.find((entry) => entry?.id === "healersBlessing") : null;
  form.querySelectorAll("label.nas-healers-blessing").forEach((node) => node.remove());
  if (!source) return;

  const container = getFlagsContainer(form);
  const metamagicDropdown = container.querySelector("details.metamagic-dropdown");
  const optionsInput = ensureMetamagicOptionsInput(form);

  const currentOptions = getCurrentMetamagicOptions(dialog, form);
  const featureStates = getFeatureStatesFromOptions(currentOptions);
  const persistentValue = source.persistent ? getPersistentFeatureState(actor, source.id) : undefined;
  let initialChecked = false;
  if (featureStates[source.id] !== undefined) {
    initialChecked = featureStates[source.id] === true;
  } else if (persistentValue !== undefined) {
    initialChecked = persistentValue === true;
  } else {
    initialChecked = source.defaultEnabled === true;
  }

  const label = document.createElement("label");
  label.classList.add("checkbox");
  label.classList.add("nas-healers-blessing");
  const input = document.createElement("input");
  input.type = "checkbox";
  input.dataset.featureId = source.id;
  input.checked = Boolean(initialChecked);
  label.textContent = ` ${source.label}${formatLimitedUsesSuffix(source)} `;
  label.insertBefore(input, label.firstChild);

  if (metamagicDropdown?.parentElement === container) {
    container.insertBefore(label, metamagicDropdown);
  } else {
    container.appendChild(label);
  }

  const nextStates = {
    ...featureStates,
    [source.id]: Boolean(initialChecked)
  };
  const nextOptions = setFeatureStatesInOptions(currentOptions, nextStates);
  DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, nextOptions);
  updateMetamagicOptionsInput(optionsInput, nextOptions);

  input.addEventListener("change", () => {
    const latestOptions = getCurrentMetamagicOptions(dialog, form);
    const latestStates = {
      ...getFeatureStatesFromOptions(latestOptions),
      [source.id]: input.checked
    };
    const updatedOptions = setFeatureStatesInOptions(latestOptions, latestStates);
    DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, updatedOptions);
    updateMetamagicOptionsInput(optionsInput, updatedOptions);
  });

  form.addEventListener("submit", () => {
    const latestOptions = getCurrentMetamagicOptions(dialog, form);
    commitPersistentFeatureStatesFromOptions(actor, latestOptions);
  }, { once: true });

  requestAttackDialogAutoLayout(dialog);
}

function renderIntenseCelebrationStandaloneControl(dialog, form, actor, sources) {
  if (!form || !actor) return;
  const source = Array.isArray(sources) ? sources.find((entry) => entry?.id === "intenseCelebration") : null;
  form.querySelectorAll("label.nas-intense-celebration").forEach((node) => node.remove());
  if (!source) return;

  const container = getFlagsContainer(form);
  const metamagicDropdown = container.querySelector("details.metamagic-dropdown");
  const optionsInput = ensureMetamagicOptionsInput(form);

  const currentOptions = getCurrentMetamagicOptions(dialog, form);
  const featureStates = getFeatureStatesFromOptions(currentOptions);
  const persistentValue = source.persistent ? getPersistentFeatureState(actor, source.id) : undefined;
  let initialChecked = false;
  if (featureStates[source.id] !== undefined) {
    initialChecked = featureStates[source.id] === true;
  } else if (persistentValue !== undefined) {
    initialChecked = persistentValue === true;
  } else {
    initialChecked = source.defaultEnabled === true;
  }

  const label = document.createElement("label");
  label.classList.add("checkbox");
  label.classList.add("nas-intense-celebration");
  const input = document.createElement("input");
  input.type = "checkbox";
  input.dataset.featureId = source.id;
  input.checked = Boolean(initialChecked);
  label.textContent = ` ${source.label}${formatLimitedUsesSuffix(source)} `;
  label.insertBefore(input, label.firstChild);

  if (metamagicDropdown?.parentElement === container) {
    container.insertBefore(label, metamagicDropdown);
  } else {
    container.appendChild(label);
  }

  const nextStates = {
    ...featureStates,
    [source.id]: Boolean(initialChecked)
  };
  const nextOptions = setFeatureStatesInOptions(currentOptions, nextStates);
  DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, nextOptions);
  updateMetamagicOptionsInput(optionsInput, nextOptions);

  input.addEventListener("change", () => {
    const latestOptions = getCurrentMetamagicOptions(dialog, form);
    const latestStates = {
      ...getFeatureStatesFromOptions(latestOptions),
      [source.id]: input.checked
    };
    const updatedOptions = setFeatureStatesInOptions(latestOptions, latestStates);
    DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, updatedOptions);
    updateMetamagicOptionsInput(optionsInput, updatedOptions);
  });

  form.addEventListener("submit", () => {
    const latestOptions = getCurrentMetamagicOptions(dialog, form);
    commitPersistentFeatureStatesFromOptions(actor, latestOptions);
  }, { once: true });

  requestAttackDialogAutoLayout(dialog);
}

function renderNaniteBloodlineArcanaStandaloneControl(dialog, form, actor, sources) {
  if (!form || !actor) return;
  const source = Array.isArray(sources) ? sources.find((entry) => entry?.id === "naniteBloodlineArcana") : null;
  form.querySelectorAll("label.nas-nanite-bloodline-arcana").forEach((node) => node.remove());
  if (!source) return;
  if (!isNaniteBloodlineArcanaEligibleForDialog(dialog)) return;

  const container = getFlagsContainer(form);
  const metamagicDropdown = container.querySelector("details.metamagic-dropdown");
  const optionsInput = ensureMetamagicOptionsInput(form);

  const currentOptions = getCurrentMetamagicOptions(dialog, form);
  const featureStates = getFeatureStatesFromOptions(currentOptions);
  const persistentValue = source.persistent ? getPersistentFeatureState(actor, source.id) : undefined;
  let initialChecked = false;
  if (featureStates[source.id] !== undefined) {
    initialChecked = featureStates[source.id] === true;
  } else if (persistentValue !== undefined) {
    initialChecked = persistentValue === true;
  } else {
    initialChecked = source.defaultEnabled === true;
  }

  const label = document.createElement("label");
  label.classList.add("checkbox");
  label.classList.add("nas-nanite-bloodline-arcana");
  const input = document.createElement("input");
  input.type = "checkbox";
  input.dataset.featureId = source.id;
  input.checked = Boolean(initialChecked);
  label.textContent = ` ${source.label}${formatLimitedUsesSuffix(source)} `;
  label.insertBefore(input, label.firstChild);

  if (metamagicDropdown?.parentElement === container) {
    container.insertBefore(label, metamagicDropdown);
  } else {
    container.appendChild(label);
  }

  const nextStates = {
    ...featureStates,
    [source.id]: Boolean(initialChecked)
  };
  const nextOptions = setFeatureStatesInOptions(currentOptions, nextStates);
  DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, nextOptions);
  updateMetamagicOptionsInput(optionsInput, nextOptions);

  input.addEventListener("change", () => {
    const latestOptions = getCurrentMetamagicOptions(dialog, form);
    const latestStates = {
      ...getFeatureStatesFromOptions(latestOptions),
      [source.id]: input.checked
    };
    const updatedOptions = setFeatureStatesInOptions(latestOptions, latestStates);
    DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, updatedOptions);
    updateMetamagicOptionsInput(optionsInput, updatedOptions);
  });

  form.addEventListener("submit", () => {
    const latestOptions = getCurrentMetamagicOptions(dialog, form);
    commitPersistentFeatureStatesFromOptions(actor, latestOptions);
  }, { once: true });

  requestAttackDialogAutoLayout(dialog);
}

function renderOneBodyTwoMindsStandaloneControl(dialog, form, actor, sources) {
  if (!form || !actor) return;
  const source = Array.isArray(sources) ? sources.find((entry) => entry?.id === "oneBodyTwoMinds") : null;
  form.querySelectorAll("label.nas-one-body-two-minds").forEach((node) => node.remove());
  if (!source) return;

  const container = getFlagsContainer(form);
  const metamagicDropdown = container.querySelector("details.metamagic-dropdown");
  const optionsInput = ensureMetamagicOptionsInput(form);

  const currentOptions = getCurrentMetamagicOptions(dialog, form);
  const featureStates = getFeatureStatesFromOptions(currentOptions);
  const persistentValue = source.persistent ? getPersistentFeatureState(actor, source.id) : undefined;
  let initialChecked = false;
  if (featureStates[source.id] !== undefined) {
    initialChecked = featureStates[source.id] === true;
  } else if (persistentValue !== undefined) {
    initialChecked = persistentValue === true;
  } else {
    initialChecked = source.defaultEnabled === true;
  }

  const label = document.createElement("label");
  label.classList.add("checkbox");
  label.classList.add("nas-one-body-two-minds");
  const input = document.createElement("input");
  input.type = "checkbox";
  input.dataset.featureId = source.id;
  input.checked = Boolean(initialChecked);
  label.textContent = ` ${source.label}${formatLimitedUsesSuffix(source)} `;
  label.insertBefore(input, label.firstChild);

  if (metamagicDropdown?.parentElement === container) {
    container.insertBefore(label, metamagicDropdown);
  } else {
    container.appendChild(label);
  }

  const nextStates = {
    ...featureStates,
    [source.id]: Boolean(initialChecked)
  };
  const nextOptions = setFeatureStatesInOptions(currentOptions, nextStates);
  DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, nextOptions);
  updateMetamagicOptionsInput(optionsInput, nextOptions);

  input.addEventListener("change", () => {
    const latestOptions = getCurrentMetamagicOptions(dialog, form);
    const latestStates = {
      ...getFeatureStatesFromOptions(latestOptions),
      [source.id]: input.checked
    };
    const updatedOptions = setFeatureStatesInOptions(latestOptions, latestStates);
    DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, updatedOptions);
    updateMetamagicOptionsInput(optionsInput, updatedOptions);
  });

  form.addEventListener("submit", () => {
    const latestOptions = getCurrentMetamagicOptions(dialog, form);
    commitPersistentFeatureStatesFromOptions(actor, latestOptions);
  }, { once: true });

  requestAttackDialogAutoLayout(dialog);
}

function getMimicMetamagicSupportedNames() {
  const names = [...Object.keys(METAMAGIC_SLOT_COSTS), "Heighten Spell"];
  return Array.from(new Set(names)).sort((a, b) => a.localeCompare(b));
}

function setPeerlessSpeedOption(options, payload = null) {
  const next = { ...(options ?? {}) };
  if (!payload) {
    delete next.peerlessSpeed;
    return next;
  }
  next.peerlessSpeed = payload;
  return next;
}

function getEligiblePeerlessSpeedChoices(dialog) {
  const pseudoSources = [
    { label: "Quicken Spell", metaName: "Quicken Spell", type: "peerless" },
    { label: "Empower Spell", metaName: "Empower Spell", type: "peerless" },
    { label: "Maximize Spell", metaName: "Maximize Spell", type: "peerless" }
  ];
  const filtered = filterMetamagicSourcesForDialog(dialog, pseudoSources);
  const canonicalNames = [];
  for (const source of filtered) {
    const canonical = (getCanonicalMetamagicName(source) ?? "").toString().trim();
    if (!canonical) continue;
    if (!canonicalNames.includes(canonical)) canonicalNames.push(canonical);
  }

  const inOrder = [];
  if (canonicalNames.includes("Quicken Spell")) inOrder.push("Quicken Spell");
  if (canonicalNames.includes("Empower Spell")) inOrder.push("Empower Spell");
  if (canonicalNames.includes("Maximize Spell")) inOrder.push("Maximize Spell");
  return inOrder;
}

function promptPeerlessSpeedCastChoice(choices) {
  return new Promise((resolve) => {
    const options = (Array.isArray(choices) ? choices : [])
      .map((name) => `<option value="${name}">${name}</option>`)
      .join("");
    const content = `
      <form>
        <div class="form-group">
          <label>${game.i18n.localize("NAS.metamagic.peerlessSpeed.castPromptLabel")}</label>
          <select name="peerlessSpeedCastChoice">${options}</select>
        </div>
      </form>
    `;
    new Dialog({
      title: game.i18n.localize("NAS.metamagic.featureNames.peerlessSpeed"),
      content,
      buttons: {
        ok: {
          label: game.i18n.localize("NAS.common.buttons.apply"),
          callback: (html) => {
            const value = (html.find('select[name="peerlessSpeedCastChoice"]').val() ?? "").toString();
            resolve(value || null);
          }
        },
        cancel: {
          label: game.i18n.localize("NAS.common.buttons.cancel"),
          callback: () => resolve(null)
        }
      },
      default: "ok",
      close: () => resolve(null)
    }).render(true);
  });
}

function setSuccorFinalRevelationOption(options, payload = null) {
  const next = { ...(options ?? {}) };
  if (!payload) {
    delete next.succorFinalRevelation;
    return next;
  }
  next.succorFinalRevelation = payload;
  return next;
}

function promptSuccorFinalRevelationCastChoice(choices) {
  return new Promise((resolve) => {
    const options = (Array.isArray(choices) ? choices : [])
      .map((name) => `<option value="${name}">${name}</option>`)
      .join("");
    const content = `
      <form>
        <div class="form-group">
          <label>${localizeMetamagic("succorFinalRevelation.selectMetamagicLabel")}</label>
          <select name="succorFinalRevelationChoice">${options}</select>
        </div>
      </form>
    `;
    new Dialog({
      title: game.i18n.localize("NAS.metamagic.featureNames.succorFinalRevelation"),
      content,
      buttons: {
        ok: {
          label: game.i18n.localize("NAS.common.buttons.apply"),
          callback: (html) => {
            const value = (html.find('select[name="succorFinalRevelationChoice"]').val() ?? "").toString();
            resolve(value || null);
          }
        },
        cancel: {
          label: game.i18n.localize("NAS.common.buttons.cancel"),
          callback: () => resolve(null)
        }
      },
      default: "ok",
      close: () => resolve(null)
    }).render(true);
  });
}

function getEligibleSuccorFinalRevelationChoices(dialog) {
  const allowed = ["Enlarge Spell", "Extend Spell", "Silent Spell", "Still Spell"];
  const pseudoSources = allowed.map((name) => ({
    label: name,
    metaName: name,
    type: "succor"
  }));
  const filtered = filterMetamagicSourcesForDialog(dialog, pseudoSources);
  const eligible = [];
  for (const source of filtered) {
    const canonical = (getCanonicalMetamagicName(source) ?? "").toString().trim();
    if (!canonical) continue;
    if (allowed.includes(canonical) && !eligible.includes(canonical)) eligible.push(canonical);
  }
  return eligible;
}

function renderPeerlessSpeedStandaloneControl(dialog, form, actor, sources) {
  if (!form || !actor) return;
  const source = Array.isArray(sources) ? sources.find((entry) => entry?.id === "peerlessSpeed") : null;
  form.querySelectorAll("label.nas-peerless-speed").forEach((node) => node.remove());
  if (!source) return;
  const eligibleChoices = getEligiblePeerlessSpeedChoices(dialog);
  const hasDamageParts = hasDamageFormula(dialog);
  if (!eligibleChoices.length) return;
  if (!hasDamageParts && !eligibleChoices.includes("Quicken Spell")) return;

  const container = getFlagsContainer(form);
  const metamagicDropdown = container.querySelector("details.metamagic-dropdown");
  const optionsInput = ensureMetamagicOptionsInput(form);
  const currentOptions = getCurrentMetamagicOptions(dialog, form);
  const featureStates = getFeatureStatesFromOptions(currentOptions);
  const persistentValue = source.persistent ? getPersistentFeatureState(actor, source.id) : undefined;
  let initialChecked = false;
  if (featureStates[source.id] !== undefined) {
    initialChecked = featureStates[source.id] === true;
  } else if (persistentValue !== undefined) {
    initialChecked = persistentValue === true;
  } else {
    initialChecked = source.defaultEnabled === true;
  }

  const label = document.createElement("label");
  label.classList.add("checkbox");
  label.classList.add("nas-peerless-speed");
  const input = document.createElement("input");
  input.type = "checkbox";
  input.dataset.featureId = source.id;
  input.checked = Boolean(initialChecked);
  label.textContent = ` ${source.label}${formatLimitedUsesSuffix(source)} `;
  label.insertBefore(input, label.firstChild);

  if (metamagicDropdown?.parentElement === container) {
    container.insertBefore(label, metamagicDropdown);
  } else {
    container.appendChild(label);
  }

  const existingPeerlessChoice = (currentOptions?.peerlessSpeed?.chosenMetaName ?? "").toString().trim();
  let nextOptions = setPeerlessSpeedOption(currentOptions, null);
  const nextStates = {
    ...featureStates,
    [source.id]: Boolean(initialChecked)
  };
  if (initialChecked) {
    const initialMetaName = existingPeerlessChoice || "Quicken Spell";
    nextOptions = setPeerlessSpeedOption(nextOptions, {
      enabled: true,
      chosenMetaName: initialMetaName
    });
  }
  nextOptions = setFeatureStatesInOptions(nextOptions, nextStates);
  DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, nextOptions);
  updateMetamagicOptionsInput(optionsInput, nextOptions);

  input.addEventListener("change", async () => {
    let latestOptions = getCurrentMetamagicOptions(dialog, form);
    let latestStates = {
      ...getFeatureStatesFromOptions(latestOptions),
      [source.id]: input.checked
    };
    latestOptions = setFeatureStatesInOptions(setPeerlessSpeedOption(latestOptions, null), latestStates);
    DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, latestOptions);
    updateMetamagicOptionsInput(optionsInput, latestOptions);

    if (input.checked) {
      const choices = getEligiblePeerlessSpeedChoices(dialog);
      const hasDamageParts = hasDamageFormula(dialog);
      const empowerOrMaximizeChoices = choices.filter((name) => name === "Empower Spell" || name === "Maximize Spell");
      if (!hasDamageParts) {
        if (!choices.includes("Quicken Spell")) {
          ui.notifications.warn(game.i18n.localize("NAS.metamagic.peerlessSpeed.notApplicable"));
          input.checked = false;
          latestStates[source.id] = false;
          const updated = setFeatureStatesInOptions(setPeerlessSpeedOption(latestOptions, null), latestStates);
          DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, updated);
          updateMetamagicOptionsInput(optionsInput, updated);
          return;
        }
        const updated = setFeatureStatesInOptions(setPeerlessSpeedOption(latestOptions, {
          enabled: true,
          chosenMetaName: "Quicken Spell"
        }), latestStates);
        DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, updated);
        updateMetamagicOptionsInput(optionsInput, updated);
        return;
      }

      let promptChoices = choices;
      if (!choices.includes("Quicken Spell") && empowerOrMaximizeChoices.length > 0) {
        promptChoices = empowerOrMaximizeChoices;
      }
      if (!promptChoices.length) {
        ui.notifications.warn(game.i18n.localize("NAS.metamagic.peerlessSpeed.notApplicable"));
        input.checked = false;
        latestStates[source.id] = false;
        const updated = setFeatureStatesInOptions(setPeerlessSpeedOption(latestOptions, null), latestStates);
        DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, updated);
        updateMetamagicOptionsInput(optionsInput, updated);
        return;
      }

      let chosenMetaName = promptChoices[0] ?? null;
      if (promptChoices.length > 1) {
        chosenMetaName = await promptPeerlessSpeedCastChoice(promptChoices);
      }
      if (!chosenMetaName) {
        input.checked = false;
        latestStates[source.id] = false;
        const updated = setFeatureStatesInOptions(setPeerlessSpeedOption(latestOptions, null), latestStates);
        DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, updated);
        updateMetamagicOptionsInput(optionsInput, updated);
        return;
      }

      const updated = setFeatureStatesInOptions(setPeerlessSpeedOption(latestOptions, {
        enabled: true,
        chosenMetaName: String(chosenMetaName)
      }), latestStates);
      DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, updated);
      updateMetamagicOptionsInput(optionsInput, updated);
      return;
    }

    latestOptions = setPeerlessSpeedOption(latestOptions, null);
    const updatedOptions = setFeatureStatesInOptions(latestOptions, latestStates);
    DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, updatedOptions);
    updateMetamagicOptionsInput(optionsInput, updatedOptions);
  });

  form.addEventListener("submit", () => {
    const latestOptions = getCurrentMetamagicOptions(dialog, form);
    commitPersistentFeatureStatesFromOptions(actor, latestOptions);
  }, { once: true });

  requestAttackDialogAutoLayout(dialog);
}

function renderSuccorFinalRevelationStandaloneControl(dialog, form, actor, sources) {
  if (!form || !actor) return;
  const source = Array.isArray(sources) ? sources.find((entry) => entry?.id === "succorFinalRevelation") : null;
  form.querySelectorAll("label.nas-succor-final-revelation").forEach((node) => node.remove());
  if (!source) return;

  const container = getFlagsContainer(form);
  const metamagicDropdown = container.querySelector("details.metamagic-dropdown");
  const optionsInput = ensureMetamagicOptionsInput(form);
  const currentOptions = getCurrentMetamagicOptions(dialog, form);
  const featureStates = getFeatureStatesFromOptions(currentOptions);
  const existingSuccorChoice = (currentOptions?.succorFinalRevelation?.chosenMetaName ?? "").toString().trim();
  const succorChoices = ["Enlarge Spell", "Extend Spell", "Silent Spell", "Still Spell"];
  const eligibleSuccorChoices = getEligibleSuccorFinalRevelationChoices(dialog);
  const persistentValue = source.persistent ? getPersistentFeatureState(actor, source.id) : undefined;
  let initialChecked = false;
  if (featureStates[source.id] !== undefined) {
    initialChecked = featureStates[source.id] === true;
  } else if (persistentValue !== undefined) {
    initialChecked = persistentValue === true;
  } else {
    initialChecked = source.defaultEnabled === true;
  }
  if (initialChecked && eligibleSuccorChoices.length < 1) {
    initialChecked = false;
  }

  const label = document.createElement("label");
  label.classList.add("checkbox");
  label.classList.add("nas-succor-final-revelation");
  const input = document.createElement("input");
  input.type = "checkbox";
  input.dataset.featureId = source.id;
  input.checked = Boolean(initialChecked);
  label.textContent = ` ${source.label}${formatLimitedUsesSuffix(source)} `;
  label.insertBefore(input, label.firstChild);

  if (metamagicDropdown?.parentElement === container) {
    container.insertBefore(label, metamagicDropdown);
  } else {
    container.appendChild(label);
  }

  const nextStates = {
    ...featureStates,
    [source.id]: Boolean(initialChecked)
  };
  let nextOptions = setSuccorFinalRevelationOption(currentOptions, null);
  if (initialChecked) {
    const initialMetaName = eligibleSuccorChoices.includes(existingSuccorChoice)
      ? existingSuccorChoice
      : (eligibleSuccorChoices[0] ?? "");
    if (initialMetaName) {
      nextOptions = setSuccorFinalRevelationOption(nextOptions, {
        enabled: true,
        chosenMetaName: initialMetaName
      });
    }
  }
  nextOptions = setFeatureStatesInOptions(nextOptions, nextStates);
  DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, nextOptions);
  updateMetamagicOptionsInput(optionsInput, nextOptions);

  input.addEventListener("change", async () => {
    let latestOptions = getCurrentMetamagicOptions(dialog, form);
    let latestStates = {
      ...getFeatureStatesFromOptions(latestOptions),
      [source.id]: input.checked
    };
    if (input.checked) {
      const eligibleChoices = getEligibleSuccorFinalRevelationChoices(dialog);
      if (!eligibleChoices.length) {
        ui.notifications.warn(localizeMetamagic("succorFinalRevelation.noEligibleChoices"));
        input.checked = false;
        latestStates[source.id] = false;
        const reverted = setFeatureStatesInOptions(setSuccorFinalRevelationOption(latestOptions, null), latestStates);
        DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, reverted);
        updateMetamagicOptionsInput(optionsInput, reverted);
        return;
      }
      const chosenMetaName = await promptSuccorFinalRevelationCastChoice(eligibleChoices);
      if (!chosenMetaName) {
        input.checked = false;
        latestStates[source.id] = false;
        const reverted = setFeatureStatesInOptions(setSuccorFinalRevelationOption(latestOptions, null), latestStates);
        DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, reverted);
        updateMetamagicOptionsInput(optionsInput, reverted);
        return;
      }
      latestOptions = setSuccorFinalRevelationOption(latestOptions, {
        enabled: true,
        chosenMetaName: String(chosenMetaName)
      });
      latestOptions = setFeatureStatesInOptions(latestOptions, latestStates);
      DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, latestOptions);
      updateMetamagicOptionsInput(optionsInput, latestOptions);
    } else {
      latestOptions = setSuccorFinalRevelationOption(latestOptions, null);
      latestOptions = setFeatureStatesInOptions(latestOptions, latestStates);
      DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, latestOptions);
      updateMetamagicOptionsInput(optionsInput, latestOptions);
    }
  });

  form.addEventListener("submit", () => {
    const latestOptions = getCurrentMetamagicOptions(dialog, form);
    commitPersistentFeatureStatesFromOptions(actor, latestOptions);
  }, { once: true });

  requestAttackDialogAutoLayout(dialog);
}

function setMimicMetamagicOption(options, payload = null) {
  const next = { ...(options ?? {}) };
  if (!payload) {
    delete next.mimicMetamagic;
    return next;
  }
  next.mimicMetamagic = payload;
  return next;
}

function getMimicMetamagicChoicePrompt({ title, description, options, preselected = [], maxSelections = 0 }) {
  return new Promise((resolve) => {
    const selectedSet = new Set(Array.isArray(preselected) ? preselected : []);
    const cards = (Array.isArray(options) ? options : [])
      .map((name) => `
        <label class="checkbox" style="display:block; margin: 4px 0;">
          <input type="checkbox" name="mimicMetamagicChoice" value="${name}" ${selectedSet.has(name) ? "checked" : ""}/>
          ${name}
        </label>
      `)
      .join("");
    const content = `
      <form>
        <p style="margin: 0 0 8px 0;">${description}</p>
        <div style="max-height: 340px; overflow-y: auto; border: 1px solid #888; border-radius: 6px; padding: 8px;">
          ${cards}
        </div>
      </form>
    `;
    new Dialog({
      title,
      content,
      buttons: {
        ok: {
          label: game.i18n.localize("NAS.common.buttons.apply"),
          callback: (html) => {
            const values = html
              .find('input[name="mimicMetamagicChoice"]:checked')
              .map((_, el) => (el?.value ?? "").toString())
              .get()
              .filter(Boolean);
            resolve(values);
          }
        },
        cancel: {
          label: game.i18n.localize("NAS.common.buttons.cancel"),
          callback: () => resolve(null)
        }
      },
      default: "ok",
      close: () => resolve(null),
      render: (html) => {
        const update = () => {
          const checked = html.find('input[name="mimicMetamagicChoice"]:checked').length;
          html.find(".nas-mimic-choice-count").text(checked.toString());
        };
        html.find('input[name="mimicMetamagicChoice"]').on("change", (event) => {
          const checked = html.find('input[name="mimicMetamagicChoice"]:checked').length;
          if (maxSelections > 0 && checked > maxSelections) {
            event.currentTarget.checked = false;
          }
          update();
        });
        const prompt = html.find("p");
        if (prompt.length) {
          prompt.append(` <b><span class="nas-mimic-choice-count">${selectedSet.size}</span>/${maxSelections}</b>`);
        }
        update();
      }
    }).render(true);
  });
}

function promptMimicMetamagicCastChoice(choices) {
  return new Promise((resolve) => {
    const options = (Array.isArray(choices) ? choices : [])
      .map((name) => `<option value="${name}">${name}</option>`)
      .join("");
    const content = `
      <form>
        <div class="form-group">
          <label>${game.i18n.localize("NAS.metamagic.mimicMetamagic.castPromptLabel")}</label>
          <select name="mimicMetamagicCastChoice">${options}</select>
        </div>
      </form>
    `;
    new Dialog({
      title: game.i18n.localize("NAS.metamagic.featureNames.mimicMetamagic"),
      content,
      buttons: {
        ok: {
          label: game.i18n.localize("NAS.common.buttons.apply"),
          callback: (html) => {
            const value = (html.find('select[name="mimicMetamagicCastChoice"]').val() ?? "").toString();
            resolve(value || null);
          }
        },
        cancel: {
          label: game.i18n.localize("NAS.common.buttons.cancel"),
          callback: () => resolve(null)
        }
      },
      default: "ok",
      close: () => resolve(null)
    }).render(true);
  });
}

function getEligibleMimicMetamagicChoices(dialog, configuredChoices) {
  const configured = Array.isArray(configuredChoices) ? configuredChoices : [];
  const pseudoSources = configured.map((name) => ({
    label: name,
    metaName: name,
    type: "mimic",
  }));
  const filtered = filterMetamagicSourcesForDialog(dialog, pseudoSources);
  const eligible = [];
  for (const source of filtered) {
    const canonical = (getCanonicalMetamagicName(source) ?? "").toString().trim();
    if (!canonical) continue;
    if (!eligible.includes(canonical)) eligible.push(canonical);
  }
  return eligible;
}

async function ensureMimicMetamagicChoicesConfigured(mimicItem, config) {
  const selectedCount = Number(config?.selectedCount ?? 1);
  const requiredChoices = Math.max(2, selectedCount * 2);
  const supported = getMimicMetamagicSupportedNames();
  const existing = (config?.configuredChoices ?? []).filter((name) => supported.includes(name));
  if (existing.length >= requiredChoices) return existing.slice(0, requiredChoices);

  const selected = await getMimicMetamagicChoicePrompt({
    title: game.i18n.localize("NAS.metamagic.mimicMetamagic.setupTitle"),
    description: game.i18n.format("NAS.metamagic.mimicMetamagic.setupDescription", { count: requiredChoices }),
    options: supported,
    preselected: existing,
    maxSelections: requiredChoices
  });
  if (!selected) return null;
  if (selected.length < requiredChoices) return null;
  const nextChoices = selected.slice(0, requiredChoices);
  const updateData = {
    [`system.flags.dictionary.${MIMIC_METAMAGIC_SELECTED_FLAG}`]: String(selectedCount)
  };
  const dictionary = mimicItem?.system?.flags?.dictionary ?? {};
  const nextChoiceKeys = new Set(
    nextChoices.map((_, index) => `${MIMIC_METAMAGIC_CHOICE_FLAG_PREFIX}${index + 1}`)
  );
  for (const key of Object.keys(dictionary)) {
    if (!key.startsWith(MIMIC_METAMAGIC_CHOICE_FLAG_PREFIX)) continue;
    if (nextChoiceKeys.has(key)) continue;
    updateData[`system.flags.dictionary.-=${key}`] = null;
  }
  nextChoices.forEach((name, index) => {
    const key = `${MIMIC_METAMAGIC_CHOICE_FLAG_PREFIX}${index + 1}`;
    updateData[`system.flags.dictionary.${key}`] = String(name);
  });
  await mimicItem?.update?.(updateData);
  const persistedDictionary = mimicItem?.system?.flags?.dictionary ?? {};
  const persistedChoices = Object.entries(persistedDictionary)
    .filter(([key]) => key.startsWith(MIMIC_METAMAGIC_CHOICE_FLAG_PREFIX))
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([, value]) => (value ?? "").toString());
  return nextChoices;
}

function renderMimicMetamagicStandaloneControl(dialog, form, actor, sources) {
  if (!form || !actor) return;
  const source = Array.isArray(sources) ? sources.find((entry) => entry?.id === MIMIC_METAMAGIC_FEATURE_ID) : null;
  form.querySelectorAll("label.nas-mimic-metamagic").forEach((node) => node.remove());
  if (!source) return;

  const container = getFlagsContainer(form);
  const metamagicDropdown = container.querySelector("details.metamagic-dropdown");
  const optionsInput = ensureMetamagicOptionsInput(form);
  const currentOptions = getCurrentMetamagicOptions(dialog, form);
  const featureStates = getFeatureStatesFromOptions(currentOptions);
  const persistentValue = source.persistent ? getPersistentFeatureState(actor, source.id) : undefined;
  let initialChecked = false;
  if (featureStates[source.id] !== undefined) {
    initialChecked = featureStates[source.id] === true;
  } else if (persistentValue !== undefined) {
    initialChecked = persistentValue === true;
  } else {
    initialChecked = source.defaultEnabled === true;
  }

  const label = document.createElement("label");
  label.classList.add("checkbox");
  label.classList.add("nas-mimic-metamagic");
  const input = document.createElement("input");
  input.type = "checkbox";
  input.dataset.featureId = source.id;
  input.checked = Boolean(initialChecked);
  label.textContent = ` ${source.label}${formatLimitedUsesSuffix(source)} `;
  label.insertBefore(input, label.firstChild);

  if (metamagicDropdown?.parentElement === container) {
    container.insertBefore(label, metamagicDropdown);
  } else {
    container.appendChild(label);
  }

  const nextStates = {
    ...featureStates,
    [source.id]: Boolean(initialChecked)
  };
  const nextOptions = setFeatureStatesInOptions(currentOptions, nextStates);
  DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, nextOptions);
  updateMetamagicOptionsInput(optionsInput, nextOptions);

  input.addEventListener("change", async () => {
    let latestOptions = getCurrentMetamagicOptions(dialog, form);
    let latestStates = {
      ...getFeatureStatesFromOptions(latestOptions),
      [source.id]: input.checked
    };

    // Persist the checkbox state immediately so async setup/update rerenders
    // keep the current checked value instead of restoring stale state.
    latestOptions = setFeatureStatesInOptions(latestOptions, latestStates);
    DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, latestOptions);
    updateMetamagicOptionsInput(optionsInput, latestOptions);

    if (input.checked) {
      const item = dialog?.action?.item ?? null;
      const mimicState = await getPsychicMimicMetamagicState(actor, item);
      if (!mimicState?.eligible || !mimicState?.mimicMetamagicItem) {
        ui.notifications.warn(game.i18n.localize("NAS.metamagic.mimicMetamagic.notEligible"));
        input.checked = false;
        latestStates[source.id] = false;
        latestOptions = setMimicMetamagicOption(latestOptions, null);
        const updated = setFeatureStatesInOptions(latestOptions, latestStates);
        DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, updated);
        updateMetamagicOptionsInput(optionsInput, updated);
        return;
      }

      const freshConfig = await getMimicMetamagicFlagState(mimicState.mimicMetamagicItem);
      const configuredChoices = await ensureMimicMetamagicChoicesConfigured(mimicState.mimicMetamagicItem, freshConfig);
      if (!configuredChoices?.length) {
        ui.notifications.warn(game.i18n.localize("NAS.metamagic.mimicMetamagic.incompleteSetup"));
        input.checked = false;
        latestStates[source.id] = false;
        latestOptions = setMimicMetamagicOption(latestOptions, null);
        const updated = setFeatureStatesInOptions(latestOptions, latestStates);
        DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, updated);
        updateMetamagicOptionsInput(optionsInput, updated);
        return;
      }

      const eligibleChoices = getEligibleMimicMetamagicChoices(dialog, configuredChoices);
      if (!eligibleChoices.length) {
        ui.notifications.warn(game.i18n.localize("NAS.metamagic.mimicMetamagic.noEligibleChoices"));
        input.checked = false;
        latestStates[source.id] = false;
        latestOptions = setMimicMetamagicOption(latestOptions, null);
        const updated = setFeatureStatesInOptions(latestOptions, latestStates);
        DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, updated);
        updateMetamagicOptionsInput(optionsInput, updated);
        return;
      }

      const chosenMetaName = await promptMimicMetamagicCastChoice(eligibleChoices);
      if (!chosenMetaName) {
        input.checked = false;
        latestStates[source.id] = false;
        latestOptions = setMimicMetamagicOption(latestOptions, null);
        const updated = setFeatureStatesInOptions(latestOptions, latestStates);
        DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, updated);
        updateMetamagicOptionsInput(optionsInput, updated);
        return;
      }

      let heightenSpellLevel = 0;
      let reachSpellSteps = 0;
      if (chosenMetaName === "Heighten Spell") {
        const baseLevel = getSpellBaseLevel(dialog);
        const levelChoices = getHeightenSpellLevelChoices(baseLevel);
        if (!levelChoices.length) {
          ui.notifications.warn(game.i18n.localize("NAS.metamagic.mimicMetamagic.heightenNoValidLevels"));
          input.checked = false;
          latestStates[source.id] = false;
          latestOptions = setMimicMetamagicOption(latestOptions, null);
          const updated = setFeatureStatesInOptions(latestOptions, latestStates);
          DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, updated);
          updateMetamagicOptionsInput(optionsInput, updated);
          return;
        }
        const pickedLevel = await promptHeightenSpellLevel(levelChoices, baseLevel);
        if (pickedLevel == null) {
          input.checked = false;
          latestStates[source.id] = false;
          latestOptions = setMimicMetamagicOption(latestOptions, null);
          const updated = setFeatureStatesInOptions(latestOptions, latestStates);
          DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, updated);
          updateMetamagicOptionsInput(optionsInput, updated);
          return;
        }
        heightenSpellLevel = Number(pickedLevel);
      }
      if (chosenMetaName === "Reach Spell") {
        const rangeUnits = dialog?.action?.range?.units ?? "touch";
        const stepChoices = getReachSpellStepChoices(rangeUnits);
        if (!stepChoices.length) {
          input.checked = false;
          latestStates[source.id] = false;
          latestOptions = setMimicMetamagicOption(latestOptions, null);
          const updated = setFeatureStatesInOptions(latestOptions, latestStates);
          DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, updated);
          updateMetamagicOptionsInput(optionsInput, updated);
          return;
        }
        const steps = await promptReachSpellSteps(stepChoices);
        if (steps == null) {
          input.checked = false;
          latestStates[source.id] = false;
          latestOptions = setMimicMetamagicOption(latestOptions, null);
          const updated = setFeatureStatesInOptions(latestOptions, latestStates);
          DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, updated);
          updateMetamagicOptionsInput(optionsInput, updated);
          return;
        }
        reachSpellSteps = Number(steps);
      }
      if (chosenMetaName === "Intensified Spell") {
        const cl = getSpellCasterLevel(dialog);
        const damageParts = getSpellDamageParts(dialog);
        if (!canIntensifyAnyDamagePart(damageParts, cl)) {
          ui.notifications.warn(localizeMetamagic("warnings.intensifiedCannotApply"));
          input.checked = false;
          latestStates[source.id] = false;
          latestOptions = setMimicMetamagicOption(latestOptions, null);
          const updated = setFeatureStatesInOptions(latestOptions, latestStates);
          DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, updated);
          updateMetamagicOptionsInput(optionsInput, updated);
          return;
        }
      }

      latestOptions = setMimicMetamagicOption(latestOptions, {
        enabled: true,
        chosenMetaName,
        heightenSpellLevel: Number.isFinite(heightenSpellLevel) ? heightenSpellLevel : 0,
        reachSpellSteps: Number.isFinite(reachSpellSteps) ? reachSpellSteps : 0
      });
      const updated = setFeatureStatesInOptions(latestOptions, latestStates);
      const liveRoot = elementFromHtmlLike(dialog?.element);
      const liveForm = liveRoot?.querySelector?.("form") ?? null;
      const liveOptionsInput = liveForm?.querySelector?.(`input[name="${METAMAGIC_OPTIONS_KEY}"]`) ?? null;
      DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, updated);
      if (liveOptionsInput) {
        updateMetamagicOptionsInput(liveOptionsInput, updated);
      }
      updateMetamagicOptionsInput(optionsInput, updated);
      return;
    }

    latestOptions = setMimicMetamagicOption(latestOptions, null);
    const updatedOptions = setFeatureStatesInOptions(latestOptions, latestStates);
    DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, updatedOptions);
    updateMetamagicOptionsInput(optionsInput, updatedOptions);
  });

  form.addEventListener("submit", () => {
    const latestOptions = getCurrentMetamagicOptions(dialog, form);
    commitPersistentFeatureStatesFromOptions(actor, latestOptions);
  }, { once: true });

  requestAttackDialogAutoLayout(dialog);
}

function updateMetamagicCheckboxOptions(dialog, listContainer, dataInput, optionsInput, metamagicSources, options = {}) {
  const storedSelections = DialogStateTracker.get(dialog.appId, METAMAGIC_SELECT_KEY) || [];
  const storedOptions = DialogStateTracker.get(dialog.appId, METAMAGIC_OPTIONS_KEY) || {};
  const infoContainer = options.infoContainer ?? null;
  const featureContainer = options.featureContainer ?? null;
  const featureSources = Array.isArray(options.featureSources) ? options.featureSources : [];
  let maskFocusSource = options.maskFocusSource ?? null;
  const actor = options.actor ?? null;
  const form = options.form ?? null;
  const actionItem = options.actionItem ?? null;
  const baseSpellLevel = Number(options.baseSpellLevel ?? 0);
  const baseSaveDc = Number.isFinite(Number(options.baseSaveDc)) ? Number(options.baseSaveDc) : null;
  const metamagicToggleInput = options.metamagicToggleInput ?? null;
  if (maskFocusSource && !filteredSourcesIncludeExtendSpell(metamagicSources)) {
    maskFocusSource = null;
    if (
      form
      && readFeatOptionsPayload(dialog, form)[MASK_FOCUS_ID]?.enabled === true
    ) {
      patchFeatOptionsMaskFocus(dialog, form, false);
    }
  }
  if (maskFocusSource) {
    ensureFeatOptionsInput(form, dialog);
  }
  listContainer.replaceChildren();

  metamagicSources
    .slice()
    .sort((a, b) => a.label.localeCompare(b.label))
    .forEach((source) => {
      const labelElement = document.createElement('label');
      labelElement.classList.add('checkbox');

      const input = document.createElement('input');
      input.type = 'checkbox';
      input.value = source.item?.uuid ?? source.item?.id ?? source.label;
      input.dataset.type = source.type;
      input.dataset.metaName = source.metaName ?? source.label;
      input.checked = storedSelections.includes(input.value);

      input.addEventListener('change', () => {
        const checkedIds = Array.from(
          listContainer.querySelectorAll('input[type="checkbox"]:checked')
        ).map((checkbox) => checkbox.value);
        DialogStateTracker.set(dialog.appId, METAMAGIC_SELECT_KEY, checkedIds);
        updateMetamagicNames(dataInput, true, listContainer);
        const currentOptions = DialogStateTracker.get(dialog.appId, METAMAGIC_OPTIONS_KEY) || {};
        updateMetamagicOptionsInput(optionsInput, currentOptions);
        updateMetamagicInfoSection(infoContainer, {
          baseSpellLevel,
          baseSaveDc,
          metamagicEnabled: metamagicToggleInput?.checked !== false,
          options: currentOptions,
          featOptions: getCurrentFeatOptions(dialog, form),
          actionItem,
          selectedMetaNames: getSelectedMetaNames(listContainer),
          dialog
        });
      });

      if (input.dataset.metaName === "Reach Spell") {
        input.addEventListener('change', async () => {
          if (!input.checked) return;
          const rangeUnits = dialog?.action?.range?.units ?? "touch";
          const stepChoices = getReachSpellStepChoices(rangeUnits);
          if (!stepChoices.length) return;
          const steps = await promptReachSpellSteps(stepChoices);
          if (steps == null) {
            input.checked = false;
            input.dispatchEvent(new Event('change'));
            return;
          }
          const nextOptions = { ...(DialogStateTracker.get(dialog.appId, METAMAGIC_OPTIONS_KEY) || {}) };
          nextOptions.reachSpellSteps = steps;
          DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, nextOptions);
          updateMetamagicOptionsInput(optionsInput, nextOptions);
          updateMetamagicInfoSection(infoContainer, {
            baseSpellLevel,
            baseSaveDc,
            metamagicEnabled: metamagicToggleInput?.checked !== false,
            options: nextOptions,
            featOptions: getCurrentFeatOptions(dialog, form),
            actionItem,
            selectedMetaNames: getSelectedMetaNames(listContainer),
            dialog
          });
        });
      }

      if (input.dataset.metaName === "Heighten Spell") {
        input.addEventListener('change', async () => {
          if (!input.checked) {
            const nextOptions = { ...(DialogStateTracker.get(dialog.appId, METAMAGIC_OPTIONS_KEY) || {}) };
            delete nextOptions.heightenSpellLevel;
            DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, nextOptions);
            updateMetamagicOptionsInput(optionsInput, nextOptions);
            updateMetamagicInfoSection(infoContainer, {
              baseSpellLevel,
              baseSaveDc,
              metamagicEnabled: metamagicToggleInput?.checked !== false,
              options: nextOptions,
              featOptions: getCurrentFeatOptions(dialog, form),
              actionItem,
              selectedMetaNames: getSelectedMetaNames(listContainer),
              dialog
            });
            return;
          }
          const baseLevel = getSpellBaseLevel(dialog);
          const levelChoices = getHeightenSpellLevelChoices(baseLevel);
          if (!levelChoices.length) return;
          const level = await promptHeightenSpellLevel(levelChoices, baseLevel);
          if (level == null) {
            input.checked = false;
            input.dispatchEvent(new Event('change'));
            return;
          }
          const nextOptions = { ...(DialogStateTracker.get(dialog.appId, METAMAGIC_OPTIONS_KEY) || {}) };
          nextOptions.heightenSpellLevel = level;
          DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, nextOptions);
          updateMetamagicOptionsInput(optionsInput, nextOptions);
          updateMetamagicInfoSection(infoContainer, {
            baseSpellLevel,
            baseSaveDc,
            metamagicEnabled: metamagicToggleInput?.checked !== false,
            options: nextOptions,
            featOptions: getCurrentFeatOptions(dialog, form),
            actionItem,
            selectedMetaNames: getSelectedMetaNames(listContainer),
            dialog
          });
        });
      }

      labelElement.textContent = ` ${source.label} `;
      labelElement.insertBefore(input, labelElement.firstChild);
      listContainer.appendChild(labelElement);
    });

  updateMetamagicNames(dataInput, true, listContainer);
  updateMetamagicOptionsInput(optionsInput, storedOptions);
  renderFeatureSection(
    dialog,
    form,
    featureContainer,
    optionsInput,
    actor,
    featureSources,
    listContainer,
    infoContainer,
    actionItem,
    baseSpellLevel,
    baseSaveDc,
    maskFocusSource,
    metamagicToggleInput
  );
  updateMetamagicInfoSection(infoContainer, {
    baseSpellLevel,
    baseSaveDc,
    metamagicEnabled: metamagicToggleInput?.checked !== false,
    options: DialogStateTracker.get(dialog.appId, METAMAGIC_OPTIONS_KEY) || storedOptions,
    featOptions: getCurrentFeatOptions(dialog, form),
    actionItem,
    selectedMetaNames: getSelectedMetaNames(listContainer),
    dialog
  });
  requestAttackDialogAutoLayout(dialog);
}

function getReachSpellStepChoices(units) {
  const order = ["touch", "close", "medium", "long"];
  const normalized = units?.toString?.().toLowerCase() ?? "";
  const startIndex = order.indexOf(normalized);
  if (startIndex === -1) return [];
  const maxSteps = order.length - 1 - startIndex;
  return Array.from({ length: maxSteps }, (_, idx) => {
    const step = idx + 1;
    return {
      step,
      from: capitalizeRangeLabel(order[startIndex]),
      to: capitalizeRangeLabel(order[startIndex + step]),
    };
  });
}

function capitalizeRangeLabel(value) {
  if (!value) return "";
  return value.charAt(0).toUpperCase() + value.slice(1);
}

function promptReachSpellSteps(stepChoices) {
  return new Promise((resolve) => {
    const options = stepChoices
      .map((choice) => `<option value="${choice.step}">${choice.from} → ${choice.to}</option>`)
      .join("");
    const content = `
      <form>
        <div class="form-group">
          <label>${localizeMetamagic("reachSpell.rangeStepsLabel")}</label>
          <select name="reachSteps">${options}</select>
        </div>
      </form>
    `;
    new Dialog({
      title: localizeMetamagic("reachSpell.title"),
      content,
      buttons: {
        ok: {
          label: game.i18n.localize("NAS.common.buttons.apply"),
          callback: (html) => {
            const value = Number(html.find('select[name="reachSteps"]').val());
            resolve(Number.isFinite(value) ? value : null);
          },
        },
        cancel: {
          label: game.i18n.localize("NAS.common.buttons.cancel"),
          callback: () => resolve(null),
        },
      },
      default: "ok",
      close: () => resolve(null),
    }).render(true);
  });
}

function getHeightenSpellLevelChoices(baseLevel) {
  const start = Number(baseLevel ?? 0);
  if (!Number.isFinite(start) || start >= 9) return [];
  return Array.from({ length: 9 - start }, (_, index) => {
    const level = start + index + 1;
    return { level, label: formatMetamagic("heightenSpell.levelLabel", { level }) };
  });
}

function promptHeightenSpellLevel(levelChoices, baseLevel) {
  return new Promise((resolve) => {
    const options = levelChoices
      .map((choice) => `<option value="${choice.level}">${choice.label}</option>`)
      .join("");
    const content = `
      <form>
        <p style="margin: 0 0 8px 0;">${formatMetamagic("heightenSpell.currentLevel", { level: `<b>${baseLevel}</b>` })}</p>
        <div class="form-group">
          <label>${localizeMetamagic("heightenSpell.heightenTo")}</label>
          <select name="heightenLevel">${options}</select>
        </div>
      </form>
    `;
    new Dialog({
      title: localizeMetamagic("heightenSpell.title"),
      content,
      buttons: {
        ok: {
          label: game.i18n.localize("NAS.common.buttons.apply"),
          callback: (html) => {
            const value = Number(html.find('select[name="heightenLevel"]').val());
            resolve(Number.isFinite(value) ? value : null);
          },
        },
        cancel: {
          label: game.i18n.localize("NAS.common.buttons.cancel"),
          callback: () => resolve(null),
        },
      },
      default: "ok",
      close: () => resolve(null),
    }).render(true);
  });
}

async function refreshMetamagicControls(dialog, form, actor, classFeatureSources = []) {
  const item = dialog?.action?.item ?? null;
  let maskFocusSource = null;
  if (getActionItemType(dialog) === "spell" && item && !isSpellLikeSpellbookItem(item)) {
    maskFocusSource = await getMaskFocusSource(actor);
  }
  const metamagicSources = await getAvailableMetamagicSources(actor, { resolveEnglishNames: true });
  const filteredSources = filterMetamagicSourcesForDialog(dialog, metamagicSources);
  if (!filteredSources.length) return;

  const maskFocusForMetamagicUi =
    maskFocusSource && filteredSourcesIncludeExtendSpell(filteredSources) ? maskFocusSource : null;

  const listContainer = form.querySelector('.metamagic-options');
  const dataInput = form.querySelector(`input[name="${METAMAGIC_NAMES_KEY}"]`);
  const optionsInput = form.querySelector(`input[name="${METAMAGIC_OPTIONS_KEY}"]`);
  if (listContainer && dataInput && optionsInput) {
    if (maskFocusForMetamagicUi) {
      ensureFeatOptionsInput(form, dialog);
    }
    const infoContainer = form.querySelector('.nas-metamagic-info');
    const featureContainer = form.querySelector('.nas-metamagic-features');
    const metamagicToggleInput = form.querySelector(`input[name="${METAMAGIC_FORM_KEY}"]`);
    updateMetamagicCheckboxOptions(dialog, listContainer, dataInput, optionsInput, filteredSources, {
      actor,
      form,
      actionItem: item,
      baseSpellLevel: getSpellBaseLevel(dialog),
      baseSaveDc: getDialogSaveDc(dialog),
      featureContainer,
      featureSources: classFeatureSources,
      maskFocusSource: maskFocusForMetamagicUi,
      infoContainer,
      metamagicToggleInput
    });
    requestAttackDialogAutoLayout(dialog);
    return;
  }

  renderMetamagicControls(dialog, form, filteredSources, {
    actor,
    actionItem: item,
    baseSpellLevel: getSpellBaseLevel(dialog),
    baseSaveDc: getDialogSaveDc(dialog),
    classFeatureSources,
    maskFocusSource
  });
}
