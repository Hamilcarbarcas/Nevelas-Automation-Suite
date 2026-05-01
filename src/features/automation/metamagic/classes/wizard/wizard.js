import { resolveEnglishName } from "../../../utils/compendiumNameResolver.js";

export const ARCANE_RESERVOIR_ARCANIST_COMPENDIUM_SOURCE =
  "Compendium.pf1.class-abilities.Item.CtDtLshBC8pc64JV";
export const ARCANE_RESERVOIR_EXPLOITER_COMPENDIUM_SOURCE =
  "Compendium.pf1e-archetypes.pf-arch-features.Item.tWiV2EJuWxUHBFCd";

const UNIVERSALIST_SCHOOL_COMPENDIUM_SOURCE = "Compendium.pf1.class-abilities.Item.tSmRmVuqYPDTyVZY";
const UNIVERSALIST_SCHOOL_ENGLISH_NAME = "Universalist School";

const METAMAGIC_MASTERY_COMPENDIUM_SOURCE = "Compendium.pf1.class-abilities.Item.L4sgpJ5DxgYddhem";
const METAMAGIC_MASTERY_ENGLISH_NAME = "Metamagic Mastery";

export const METAMAGIC_MASTERY_FEATURE_ID = "metamagicMastery";

const WIZARD_MIN_LEVEL_FOR_MASTERY = 8;

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

function getWizardClassKeys(actor) {
  const classes = actor?.classes ?? {};
  const keys = new Set();
  for (const [classKey, classData] of Object.entries(classes)) {
    const candidates = [classKey, classData?.tag, classData?.name, classData?._id, classData?.id]
      .map(normalizeKey)
      .filter(Boolean);
    if (candidates.some((v) => v.includes("wizard") || v === "wiz")) {
      keys.add(normalizeKey(classKey));
      candidates.forEach((v) => keys.add(v));
    }
  }
  if (!keys.size) {
    keys.add("wizard");
    keys.add("wiz");
  }
  return keys;
}

export function getWizardClassLevel(actor) {
  const classes = actor?.classes ?? {};
  const keys = getWizardClassKeys(actor);
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

export function isWizardSpellItem(actor, spellItem) {
  if (!actor || !spellItem || spellItem.type !== "spell") return false;
  const bookClass = getSpellbookClassKey(spellItem);
  if (!bookClass) return false;
  const keys = getWizardClassKeys(actor);
  if (keys.has(bookClass)) return true;
  return bookClass.includes("wiz");
}

/**
 * Highest spell level this spellbook has slots for (max > 0), for Metamagic Mastery cap checks.
 */
export function getSpellbookMaxCastableSpellLevel(actor, spellbookKey) {
  const book = actor?.system?.attributes?.spells?.spellbooks?.[spellbookKey];
  if (!book) return 0;
  let max = 0;
  for (let L = 1; L <= 9; L += 1) {
    const data = book.spells?.[`spell${L}`];
    const m = Number(data?.max ?? 0);
    if (Number.isFinite(m) && m > 0) max = L;
  }
  return max;
}

async function buildWizardMetamagicMasteryState(actor, spellItem = null) {
  if (!actor) {
    return {
      eligible: false,
      universalistItem: null,
      metamagicMasteryItem: null,
      wizardLevel: 0,
      isWizardSpell: false
    };
  }

  const [universalistItem, metamagicMasteryItem] = await Promise.all([
    findActorFeat(actor, {
      compendiumSource: UNIVERSALIST_SCHOOL_COMPENDIUM_SOURCE,
      englishName: UNIVERSALIST_SCHOOL_ENGLISH_NAME
    }),
    findActorFeat(actor, {
      compendiumSource: METAMAGIC_MASTERY_COMPENDIUM_SOURCE,
      englishName: METAMAGIC_MASTERY_ENGLISH_NAME
    })
  ]);

  const wizardLevel = getWizardClassLevel(actor);
  const isWizardSpell = spellItem ? isWizardSpellItem(actor, spellItem) : true;

  const eligible =
    Boolean(universalistItem)
    && Boolean(metamagicMasteryItem)
    && wizardLevel >= WIZARD_MIN_LEVEL_FOR_MASTERY
    && isWizardSpell;

  return {
    eligible,
    universalistItem,
    metamagicMasteryItem,
    wizardLevel,
    isWizardSpell
  };
}

function getMasteryUseSnapshot(item) {
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

export async function getWizardMetamagicMasteryState(actor, spellItem = null) {
  const base = await buildWizardMetamagicMasteryState(actor, spellItem);
  const snap = getMasteryUseSnapshot(base.metamagicMasteryItem);
  return {
    ...base,
    hasUsesData: snap.hasUsesData,
    usesRemaining: snap.remaining,
    usesMax: snap.max,
    hasRemainingUses: snap.hasUsesData && snap.remaining > 0
  };
}

export async function getWizardMetamagicFeatureSources(actor, spellItem) {
  const state = await getWizardMetamagicMasteryState(actor, spellItem);
  if (!state.eligible || !state.metamagicMasteryItem) return [];

  return [
    {
      id: METAMAGIC_MASTERY_FEATURE_ID,
      label: state.metamagicMasteryItem?.name ?? METAMAGIC_MASTERY_ENGLISH_NAME,
      itemUuid: state.metamagicMasteryItem?.uuid ?? null,
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
