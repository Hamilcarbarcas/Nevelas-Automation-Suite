import { resolveEnglishName } from "../../../utils/compendiumNameResolver.js";
import { resolveMetamagicNameFromDatabase } from "../../metamagic.js";

const ORACLE_CURSE_COMPENDIUM_SOURCE = "Compendium.pf1.class-abilities.Item.kkVnlOkjhww8MoHc";
const ORACLE_CURSE_ENGLISH_NAME = "Oracle's Curse";

const SEEKER_ETERNAL_EMPEROR_COMPENDIUM_SOURCE = "Compendium.pf-content.pf-feats.Item.cA22u3PfW9FK5EeP";
const SEEKER_ETERNAL_EMPEROR_ENGLISH_NAME = "Seeker of the Eternal Emperor";
export const SEEKER_ETERNAL_EMPEROR_FEATURE_ID = "seekerOfTheEternalEmperor";
const SUCCOR_FINAL_REVELATION_COMPENDIUM_SOURCE = "";
const SUCCOR_FINAL_REVELATION_ENGLISH_NAME = "Succor Final Revelation";
const SUCCOR_MYSTERY_COMPENDIUM_SOURCE = "Compendium.pf1.class-abilities.Item.OjUWyvWC2WDjORSw";
export const SUCCOR_FINAL_REVELATION_FEATURE_ID = "succorFinalRevelation";

const DIVINATION_SCHOOL_KEYS = new Set(["div", "divination"]);
const SEEKER_REQUIRED_METAMAGIC = new Set(["Extend Spell", "Enlarge Spell"]);

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

async function findFeat(actor, { compendiumSource, englishName }) {
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

function getOracleClassKeys(actor) {
  const classes = actor?.classes ?? {};
  const keys = new Set();
  for (const [classKey, classData] of Object.entries(classes)) {
    const candidates = [classKey, classData?.tag, classData?.name, classData?._id, classData?.id]
      .map(normalizeKey)
      .filter(Boolean);
    if (candidates.some((value) => value.includes("oracle") || value === "ora")) {
      keys.add(normalizeKey(classKey));
      candidates.forEach((value) => keys.add(value));
    }
  }
  if (!keys.size) {
    keys.add("oracle");
    keys.add("ora");
  }
  return keys;
}

function getOracleClassLevel(actor) {
  const classes = actor?.classes ?? {};
  let highestLevel = 0;
  for (const [classKey, classData] of Object.entries(classes)) {
    const candidates = [classKey, classData?.tag, classData?.name]
      .map(normalizeKey)
      .filter(Boolean);
    if (!candidates.some((value) => value.includes("oracle") || value === "ora")) continue;
    const level = Number(classData?.level ?? classData?.unlevel ?? 0);
    if (Number.isFinite(level)) highestLevel = Math.max(highestLevel, level);
  }
  return highestLevel;
}

function getSpellbookClassKey(spellItem) {
  const cls = spellItem?.spellbook?.class ?? spellItem?.system?.spellbook?.class ?? "";
  return normalizeKey(cls);
}

export function isOracleSpellItem(actor, spellItem) {
  if (!actor || !spellItem || spellItem.type !== "spell") return false;
  const spellbookClass = getSpellbookClassKey(spellItem);
  if (!spellbookClass) return false;
  const classKeys = getOracleClassKeys(actor);
  if (classKeys.has(spellbookClass)) return true;
  return spellbookClass.includes("ora");
}

function isDivinationSpell(spellItem) {
  const school = normalizeKey(spellItem?.system?.school);
  return DIVINATION_SCHOOL_KEYS.has(school);
}

async function hasRequiredMetamagicFeat(actor) {
  const items = Array.from(actor?.items ?? []).filter(isFeatItem);
  if (!items.length) return false;

  for (const item of items) {
    const canonical = resolveMetamagicNameFromDatabase(item?.name ?? "");
    if (canonical && SEEKER_REQUIRED_METAMAGIC.has(canonical)) return true;
    const original = getBabeleOriginalName(item);
    const canonicalOriginal = resolveMetamagicNameFromDatabase(original ?? "");
    if (canonicalOriginal && SEEKER_REQUIRED_METAMAGIC.has(canonicalOriginal)) return true;
  }

  if (isEnglishLanguage()) return false;
  for (const item of items) {
    const resolved = await resolveEnglishName(item?.name, { documentName: "Item", deepScanMode: "off" });
    const canonical = resolveMetamagicNameFromDatabase(resolved ?? "");
    if (canonical && SEEKER_REQUIRED_METAMAGIC.has(canonical)) return true;
  }
  return false;
}

export async function getOracleSeekerFeatureSources(actor, spellItem) {
  if (!actor || !spellItem || spellItem.type !== "spell") return [];
  if (!isOracleSpellItem(actor, spellItem)) return [];
  if (!isDivinationSpell(spellItem)) return [];

  const [oracleCurseItem, seekerItem, hasMetamagicFeat] = await Promise.all([
    findFeat(actor, {
      compendiumSource: ORACLE_CURSE_COMPENDIUM_SOURCE,
      englishName: ORACLE_CURSE_ENGLISH_NAME
    }),
    findFeat(actor, {
      compendiumSource: SEEKER_ETERNAL_EMPEROR_COMPENDIUM_SOURCE,
      englishName: SEEKER_ETERNAL_EMPEROR_ENGLISH_NAME
    }),
    hasRequiredMetamagicFeat(actor)
  ]);
  if (!oracleCurseItem || !hasMetamagicFeat) return [];

  return [{
    id: SEEKER_ETERNAL_EMPEROR_FEATURE_ID,
    label: seekerItem?.name ?? SEEKER_ETERNAL_EMPEROR_ENGLISH_NAME,
    itemUuid: seekerItem?.uuid ?? null,
    limited: false,
    persistent: true,
    defaultEnabled: false,
    hasUsesData: true,
    hasRemaining: true,
    dependencies: {
      oracleCurseItemUuid: oracleCurseItem?.uuid ?? null
    },
    requiredMetamagicNames: ["Extend Spell", "Enlarge Spell"]
  }];
}

export async function getOracleSuccorFinalRevelationFeatureSources(actor, spellItem) {
  if (!actor || !spellItem || spellItem.type !== "spell") return [];
  if (!isOracleSpellItem(actor, spellItem)) return [];
  if (getOracleClassLevel(actor) < 20) return [];

  const [succorMysteryItem, succorFinalRevelationItem] = await Promise.all([
    findFeat(actor, {
      compendiumSource: SUCCOR_MYSTERY_COMPENDIUM_SOURCE,
      englishName: "Succor Mystery"
    }),
    findFeat(actor, {
      compendiumSource: SUCCOR_FINAL_REVELATION_COMPENDIUM_SOURCE,
      englishName: SUCCOR_FINAL_REVELATION_ENGLISH_NAME
    })
  ]);
  if (!succorMysteryItem || !succorFinalRevelationItem) return [];

  return [{
    id: SUCCOR_FINAL_REVELATION_FEATURE_ID,
    label: succorFinalRevelationItem?.name ?? SUCCOR_FINAL_REVELATION_ENGLISH_NAME,
    itemUuid: succorFinalRevelationItem?.uuid ?? null,
    limited: false,
    persistent: false,
    defaultEnabled: false,
    hasUsesData: true,
    hasRemaining: true,
    dependencies: {
      succorMysteryItemUuid: succorMysteryItem?.uuid ?? null
    }
  }];
}
