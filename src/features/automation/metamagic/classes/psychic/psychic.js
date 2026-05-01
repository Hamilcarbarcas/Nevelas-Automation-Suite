import { resolveEnglishName } from "../../../utils/compendiumNameResolver.js";
import {
  evaluateSpellAbilityEligibility,
  isSpellAbilityMinimumEnabled
} from "../../../spellcasting/abilityLimit.js";

const PSYCHIC_CLASS_NAME = "psychic";
const PSYCHIC_MIN_LEVEL_FOR_MIMIC = 11;

const PHRENIC_POOL_COMPENDIUM_SOURCE = "Compendium.pf1.class-abilities.Item.tFy3rxyljSq56HSg";
const PHRENIC_POOL_ENGLISH_NAME = "Phrenic Pool";
const MAJOR_AMPLIFICATIONS_COMPENDIUM_SOURCE = "Compendium.pf1.class-abilities.Item.nO6cI6JoY2THQoxZ";
const MAJOR_AMPLIFICATIONS_ENGLISH_NAME = "Major Amplifications";
const MIMIC_METAMAGIC_COMPENDIUM_SOURCE = "Compendium.pf1.class-abilities.Item.6QoaFhevY4s7zYfp";
const MIMIC_METAMAGIC_ENGLISH_NAME = "Mimic Metamagic";

export const MIMIC_METAMAGIC_FEATURE_ID = "mimicMetamagic";
export const MIMIC_METAMAGIC_SELECTED_FLAG = "MimicSelected";
export const MIMIC_METAMAGIC_CHOICE_FLAG_PREFIX = "MimicChoice";

function normalizeKey(value) {
  return (value ?? "").toString().trim().toLowerCase();
}

function isEnglishLanguage() {
  return (game?.i18n?.lang ?? "en").toLowerCase().startsWith("en");
}

function isFeatItem(item) {
  const subType = item?.subType ?? item?.system?.subType;
  return item?.type === "feat" && (subType === "feat" || subType === "classFeat");
}

function getBabeleOriginalName(item) {
  const direct = item?.flags?.babele?.originalName;
  const canUseBabele = game?.modules?.get?.("babele")?.active === true;
  const hasGetFlag = canUseBabele && typeof item?.getFlag === "function";
  let viaGetFlag = null;
  try {
    viaGetFlag = hasGetFlag ? item.getFlag("babele", "originalName") : null;
  } catch (_err) {
    viaGetFlag = null;
  }
  return direct ?? viaGetFlag ?? null;
}

function getFeatFastMatchReason(item, { compendiumSource, englishName }) {
  const source = item?._stats?.compendiumSource ?? "";
  if (source && source === compendiumSource) return "compendiumSource";
  const originalName = getBabeleOriginalName(item);
  if (originalName && originalName === englishName) return "babele.originalName";
  if (isEnglishLanguage() && (item?.name ?? "") === englishName) return "name";
  return null;
}

async function findActorFeat(actor, { compendiumSource, englishName }) {
  const items = actor?.items;
  if (!items) return null;
  const feats = Array.from(items).filter(isFeatItem);
  for (const feat of feats) {
    const reason = getFeatFastMatchReason(feat, { compendiumSource, englishName });
    if (reason) return feat;
  }
  if (isEnglishLanguage()) return null;
  for (const feat of feats) {
    const resolved = await resolveEnglishName(feat?.name, { documentName: "Item", deepScanMode: "off" });
    if ((resolved ?? "") === englishName) return feat;
  }
  return null;
}

function getPsychicClassKeys(actor) {
  const classes = actor?.classes ?? {};
  const keys = new Set();
  for (const [classKey, classData] of Object.entries(classes)) {
    const candidates = [classKey, classData?.tag, classData?.name, classData?._id, classData?.id]
      .map(normalizeKey)
      .filter(Boolean);
    if (candidates.some((v) => v.includes(PSYCHIC_CLASS_NAME) || v === "psy")) {
      keys.add(normalizeKey(classKey));
      candidates.forEach((v) => keys.add(v));
    }
  }
  if (!keys.size) {
    keys.add(PSYCHIC_CLASS_NAME);
    keys.add("psy");
  }
  return keys;
}

export function getPsychicClassLevel(actor) {
  const classes = actor?.classes ?? {};
  const keys = getPsychicClassKeys(actor);
  const levels = [];
  for (const [classKey, classData] of Object.entries(classes)) {
    const normalizedClassKey = normalizeKey(classKey);
    const tag = normalizeKey(classData?.tag);
    const name = normalizeKey(classData?.name);
    if (!keys.has(normalizedClassKey) && !keys.has(tag) && !keys.has(name)) continue;
    const levelValue = Number(classData?.level ?? 0);
    if (Number.isFinite(levelValue) && levelValue > 0) levels.push(levelValue);
  }
  if (!levels.length) return 0;
  return Math.max(...levels);
}

function getSpellbookClassKey(spellItem) {
  const cls = spellItem?.spellbook?.class ?? spellItem?.system?.spellbook?.class ?? "";
  return normalizeKey(cls);
}

export function isPsychicSpellItem(actor, spellItem) {
  if (!actor || !spellItem || spellItem.type !== "spell") return false;
  const bookClass = getSpellbookClassKey(spellItem);
  if (!bookClass) return false;
  const keys = getPsychicClassKeys(actor);
  if (keys.has(bookClass)) return true;
  return bookClass.includes("psy");
}

export function getPsychicSpellbookMaxCastableSpellLevel(actor, spellbookKey) {
  const book = actor?.system?.attributes?.spells?.spellbooks?.[spellbookKey];
  if (!book) return 0;
  const enforceAbilityMinimum = isSpellAbilityMinimumEnabled();
  let max = 0;
  for (let level = 1; level <= 9; level += 1) {
    const data = book?.spells?.[`spell${level}`];
    const maxSlots = Number(data?.max ?? 0);
    if (!Number.isFinite(maxSlots) || maxSlots <= 0) continue;
    if (data?.lowLevel === true) continue;
    if (enforceAbilityMinimum) {
      const eligibility = evaluateSpellAbilityEligibility(actor, {
        spellbookKey,
        spellLevel: level,
        honorNoAbilityLimit: true
      });
      if (!eligibility.allowed) continue;
    }
    max = level;
  }
  return max;
}

function getUseSnapshot(item) {
  const uses = item?.system?.uses;
  if (!uses?.per) {
    return { hasUsesData: false, remaining: 0, max: 0 };
  }
  const remaining = Number(uses.value ?? 0);
  const max = Number(uses.max ?? 0);
  return {
    hasUsesData: true,
    remaining: Number.isFinite(remaining) ? remaining : 0,
    max: Number.isFinite(max) ? max : 0
  };
}

function normalizeMimicMetamagicChoices(rawChoices) {
  if (!Array.isArray(rawChoices)) return [];
  const out = [];
  for (const value of rawChoices) {
    const name = (value ?? "").toString().trim();
    if (!name) continue;
    if (!out.includes(name)) out.push(name);
  }
  return out;
}

export async function getMimicMetamagicFlagState(mimicItem) {
  if (!mimicItem) {
    return {
      selectedCount: 1,
      requiredChoices: 2,
      configuredChoices: [],
      isComplete: false
    };
  }

  const dictionary = mimicItem?.system?.flags?.dictionary ?? {};
  const selectedRaw = Number(dictionary?.[MIMIC_METAMAGIC_SELECTED_FLAG] ?? 1);
  const selectedCount = Number.isFinite(selectedRaw) && selectedRaw > 0
    ? Math.max(1, Math.floor(selectedRaw))
    : 1;
  const requiredChoices = selectedCount * 2;
  const choiceEntries = Object.entries(dictionary)
    .filter(([key]) => key.startsWith(MIMIC_METAMAGIC_CHOICE_FLAG_PREFIX))
    .map(([key, value]) => {
      const indexRaw = key.slice(MIMIC_METAMAGIC_CHOICE_FLAG_PREFIX.length);
      const index = Number(indexRaw);
      return { index: Number.isFinite(index) ? index : Number.POSITIVE_INFINITY, value };
    })
    .sort((a, b) => a.index - b.index)
    .map((entry) => entry.value);
  const configuredChoices = normalizeMimicMetamagicChoices(choiceEntries);

  return {
    selectedCount,
    requiredChoices,
    configuredChoices,
    isComplete: configuredChoices.length >= requiredChoices
  };
}

export async function getPsychicMimicMetamagicState(actor, spellItem = null) {
  if (!actor) {
    return {
      eligible: false,
      psychicLevel: 0,
      isPsychicSpell: false,
      phrenicPoolItem: null,
      majorAmplificationsItem: null,
      mimicMetamagicItem: null,
      hasUsesData: false,
      usesRemaining: 0,
      usesMax: 0,
      hasRemainingUses: false,
      config: {
        selectedCount: 1,
        requiredChoices: 2,
        configuredChoices: [],
        isComplete: false
      }
    };
  }

  const [phrenicPoolItem, majorAmplificationsItem, mimicMetamagicItem] = await Promise.all([
    findActorFeat(actor, {
      compendiumSource: PHRENIC_POOL_COMPENDIUM_SOURCE,
      englishName: PHRENIC_POOL_ENGLISH_NAME
    }),
    findActorFeat(actor, {
      compendiumSource: MAJOR_AMPLIFICATIONS_COMPENDIUM_SOURCE,
      englishName: MAJOR_AMPLIFICATIONS_ENGLISH_NAME
    }),
    findActorFeat(actor, {
      compendiumSource: MIMIC_METAMAGIC_COMPENDIUM_SOURCE,
      englishName: MIMIC_METAMAGIC_ENGLISH_NAME
    })
  ]);

  const psychicLevel = getPsychicClassLevel(actor);
  const isPsychicSpell = spellItem ? isPsychicSpellItem(actor, spellItem) : true;
  const useSnap = getUseSnapshot(phrenicPoolItem);
  const config = await getMimicMetamagicFlagState(mimicMetamagicItem);

  const eligible =
    Boolean(phrenicPoolItem)
    && Boolean(majorAmplificationsItem)
    && Boolean(mimicMetamagicItem)
    && psychicLevel >= PSYCHIC_MIN_LEVEL_FOR_MIMIC
    && isPsychicSpell;

  return {
    eligible,
    psychicLevel,
    isPsychicSpell,
    phrenicPoolItem,
    majorAmplificationsItem,
    mimicMetamagicItem,
    hasUsesData: useSnap.hasUsesData,
    usesRemaining: useSnap.remaining,
    usesMax: useSnap.max,
    hasRemainingUses: useSnap.hasUsesData && useSnap.remaining > 0,
    config
  };
}

export async function getPsychicMimicMetamagicFeatureSources(actor, spellItem) {
  const state = await getPsychicMimicMetamagicState(actor, spellItem);
  if (!state.eligible || !state.mimicMetamagicItem || !state.phrenicPoolItem) return [];
  return [
    {
      id: MIMIC_METAMAGIC_FEATURE_ID,
      label: state.mimicMetamagicItem?.name ?? MIMIC_METAMAGIC_ENGLISH_NAME,
      itemUuid: state.mimicMetamagicItem?.uuid ?? null,
      limited: true,
      persistent: false,
      defaultEnabled: false,
      hasUsesData: state.hasUsesData,
      hasRemaining: state.hasRemainingUses,
      usesValue: state.usesRemaining,
      usesMax: state.usesMax
    }
  ];
}
