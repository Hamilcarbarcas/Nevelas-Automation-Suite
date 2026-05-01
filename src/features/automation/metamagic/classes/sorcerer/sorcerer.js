import { resolveEnglishName } from "../../../utils/compendiumNameResolver.js";
import { isDurationEligibleForExtendSpell } from "../../extendSpell.js";

const ARCANE_BLOODLINE_COMPENDIUM_SOURCE = "Compendium.pf1.class-abilities.Item.Y926YHNobEojovx4";
const ARCANE_BLOODLINE_ENGLISH_NAME = "Arcane Bloodline";

const METAMAGIC_ADEPT_COMPENDIUM_SOURCE = "Compendium.pf1.class-abilities.Item.EwV7Db8W6ww3rtd0";
const METAMAGIC_ADEPT_ENGLISH_NAME = "Metamagic Adept";

const ARCANE_APOTHEOSIS_COMPENDIUM_SOURCE = "Compendium.pf1.class-abilities.Item.bS39nw77bwZkH8KQ";
const ARCANE_APOTHEOSIS_ENGLISH_NAME = "Arcane Apotheosis";
const MARTYRED_BLOODLINE_COMPENDIUM_SOURCE = "Compendium.pf1.class-abilities.Item.mSZ4PrieWFBh3Xpo";
const MARTYRED_BLOODLINE_ENGLISH_NAME = "Martyred Bloodline";
const RETRIBUTION_COMPENDIUM_SOURCE = "Compendium.pf-content.pf-class-abilities.Item.qCoNRxyflZtuzJFo";
const RETRIBUTION_ENGLISH_NAME = "Retribution";
const MAESTRO_BLOODLINE_COMPENDIUM_SOURCE = "Compendium.pf1.class-abilities.Item.a0CjjkvzrCEOPyIE";
const MAESTRO_BLOODLINE_ENGLISH_NAME = "Maestro Bloodline";
const GRAND_MAESTRO_COMPENDIUM_SOURCE = "Compendium.pf1.class-abilities.Item.HSa7LF5k1v4QwYt2";
const GRAND_MAESTRO_ENGLISH_NAME = "Grand Maestro";
const NANITE_BLOODLINE_COMPENDIUM_SOURCE = "Compendium.pf1.class-abilities.Item.xcbwys8JTF22w8oW";
const NANITE_BLOODLINE_ENGLISH_NAME = "Nanite Bloodline";
const ASTRAL_BLOODLINE_COMPENDIUM_SOURCE = "Compendium.pf1.class-abilities.Item.Bv4TWEti93H40Qbf";
const ASTRAL_BLOODLINE_ENGLISH_NAME = "Astral Bloodline";
const TIMELESS_SOUL_COMPENDIUM_SOURCE = "Compendium.pf1.class-abilities.Item.sdPGUAeJkb8LuSE4";
const TIMELESS_SOUL_ENGLISH_NAME = "Timeless Soul";
const POSSESSED_BLOODLINE_COMPENDIUM_SOURCE = "Compendium.pf1.class-abilities.Item.0EcvlKOKOAurc0DS";
const POSSESSED_BLOODLINE_ENGLISH_NAME = "Possessed Bloodline";
const PEERLESS_SPEED_COMPENDIUM_SOURCE = "Compendium.pf1.class-abilities.Item.LI7UgRNGJ77FYqio";
const PEERLESS_SPEED_ENGLISH_NAME = "Peerless Speed";
const PEERLESS_SPEED_REQUIRED_LEVEL = 3;
const ONE_BODY_TWO_MINDS_COMPENDIUM_SOURCE = "Compendium.pf1.class-abilities.Item.UqUeqBlJqWUFdQwC";
const ONE_BODY_TWO_MINDS_ENGLISH_NAME = "One Body, Two Minds";
const ONE_BODY_TWO_MINDS_REQUIRED_LEVEL = 15;

const TRANSMUTATION_SCHOOL_KEYS = new Set(["trs", "transmutation", "tra"]);

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
  if (!feats.length) return null;

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

function getSorcererClassKeys(actor) {
  const classes = actor?.classes ?? {};
  const entries = Object.entries(classes);
  const keys = new Set();

  for (const [classKey, classData] of entries) {
    const candidates = [
      classKey,
      classData?.tag,
      classData?.name,
      classData?._id,
      classData?.id
    ]
      .map(normalizeKey)
      .filter(Boolean);
    if (candidates.some((value) => value.includes("sorcerer") || value === "sor")) {
      keys.add(normalizeKey(classKey));
      candidates.forEach((value) => keys.add(value));
    }
  }

  if (!keys.size) {
    keys.add("sorcerer");
    keys.add("sor");
  }
  return keys;
}

function getSorcererClassLevel(actor) {
  const classes = actor?.classes ?? {};
  const keys = getSorcererClassKeys(actor);
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

function isNaniteBloodlineEligibleSpell(spellItem, { durationOverride = null } = {}) {
  if (!spellItem || spellItem.type !== "spell") return false;
  const school = normalizeKey(spellItem?.system?.school);
  if (!TRANSMUTATION_SCHOOL_KEYS.has(school)) return false;

  const resolvedDuration =
    durationOverride ??
    spellItem?.action?.duration ??
    spellItem?.actions?.[0]?.duration ??
    spellItem?.item?.system?.actions?.[0]?.duration ??
    spellItem?.system?.actions?.[0]?.duration ??
    spellItem?.system?.duration;

  const eligible = isDurationEligibleForExtendSpell(resolvedDuration);
  return eligible;
}

export function isSorcererSpellItem(actor, spellItem) {
  if (!actor || !spellItem) return false;
  if (spellItem.type !== "spell") return false;
  const spellbookClass = getSpellbookClassKey(spellItem);
  if (!spellbookClass) return false;
  const classKeys = getSorcererClassKeys(actor);
  if (classKeys.has(spellbookClass)) return true;
  return spellbookClass.includes("sor");
}

function getUseState(item, { limited = false } = {}) {
  if (!limited) {
    return {
      limited: false,
      hasUsesData: true,
      remaining: null,
      max: null,
      hasRemaining: true
    };
  }

  const uses = item?.system?.uses;
  if (!uses || !uses.per) {
    return {
      limited: true,
      hasUsesData: false,
      remaining: 0,
      max: 0,
      hasRemaining: false
    };
  }

  const remaining = Number(uses.value ?? 0);
  const max = Number(uses.max ?? 0);
  return {
    limited: true,
    hasUsesData: true,
    remaining: Number.isFinite(remaining) ? remaining : 0,
    max: Number.isFinite(max) ? max : 0,
    hasRemaining: Number.isFinite(remaining) && remaining > 0
  };
}

async function buildSorcererArcaneState(actor, spellItem = null) {
  if (!actor) {
    return {
      actor: null,
      spellItem: spellItem ?? null,
      isSorcererSpell: false,
      sorcererLevel: 0,
      hasArcaneBloodline: false,
      hasMaestroBloodline: false,
      hasGrandMaestro: false,
      hasMetamagicAdept: false,
      hasArcaneApotheosis: false,
      arcaneBloodlineItem: null,
      maestroBloodlineItem: null,
      grandMaestroItem: null,
      metamagicAdeptItem: null,
      arcaneApotheosisItem: null
    };
  }

  const [arcaneBloodlineItem, maestroBloodlineItem, grandMaestroItem, metamagicAdeptItem, arcaneApotheosisItem] = await Promise.all([
    findFeat(actor, {
      compendiumSource: ARCANE_BLOODLINE_COMPENDIUM_SOURCE,
      englishName: ARCANE_BLOODLINE_ENGLISH_NAME
    }),
    findFeat(actor, {
      compendiumSource: MAESTRO_BLOODLINE_COMPENDIUM_SOURCE,
      englishName: MAESTRO_BLOODLINE_ENGLISH_NAME
    }),
    findFeat(actor, {
      compendiumSource: GRAND_MAESTRO_COMPENDIUM_SOURCE,
      englishName: GRAND_MAESTRO_ENGLISH_NAME
    }),
    findFeat(actor, {
      compendiumSource: METAMAGIC_ADEPT_COMPENDIUM_SOURCE,
      englishName: METAMAGIC_ADEPT_ENGLISH_NAME
    }),
    findFeat(actor, {
      compendiumSource: ARCANE_APOTHEOSIS_COMPENDIUM_SOURCE,
      englishName: ARCANE_APOTHEOSIS_ENGLISH_NAME
    })
  ]);

  const sorcererLevel = getSorcererClassLevel(actor);
  const isSorcererSpell = spellItem ? isSorcererSpellItem(actor, spellItem) : true;

  return {
    actor,
    spellItem: spellItem ?? null,
    isSorcererSpell,
    sorcererLevel,
    hasArcaneBloodline: Boolean(arcaneBloodlineItem),
    hasMaestroBloodline: Boolean(maestroBloodlineItem),
    hasGrandMaestro: Boolean(grandMaestroItem),
    hasMetamagicAdept: Boolean(metamagicAdeptItem),
    hasArcaneApotheosis: Boolean(arcaneApotheosisItem),
    arcaneBloodlineItem,
    maestroBloodlineItem,
    grandMaestroItem,
    metamagicAdeptItem,
    arcaneApotheosisItem
  };
}

export async function getSorcererArcaneMetamagicState(actor, spellItem = null) {
  const base = await buildSorcererArcaneState(actor, spellItem);
  const level = Number(base.sorcererLevel ?? 0);
  const isEligibleActor = Boolean(base.hasArcaneBloodline && level > 0);
  const canUsePassive = Boolean(base.isSorcererSpell && isEligibleActor);

  const metamagicAdeptAvailable =
    canUsePassive &&
    base.hasMetamagicAdept &&
    level >= 3 &&
    level < 20;
  const arcaneApotheosisAvailable =
    canUsePassive &&
    base.hasArcaneApotheosis &&
    level >= 20;
  const grandMaestroAvailable =
    Boolean(base.isSorcererSpell) &&
    base.hasMaestroBloodline &&
    base.hasGrandMaestro &&
    level >= 20;

  const metamagicAdeptUse = getUseState(base.metamagicAdeptItem, { limited: true });
  const arcaneApotheosisUse = getUseState(base.arcaneApotheosisItem, { limited: false });

  return {
    ...base,
    isEligibleActor,
    canUsePassive,
    metamagicAdeptAvailable,
    arcaneApotheosisAvailable,
    grandMaestroAvailable,
    metamagicAdeptUse,
    arcaneApotheosisUse
  };
}

export async function getSorcererMetamagicFeatureSources(actor, spellItem) {
  const state = await getSorcererArcaneMetamagicState(actor, spellItem);
  const sources = [];
  const [martyredBloodlineItem, retributionItem, timelessSoulSources] = await Promise.all([
    findFeat(actor, {
      compendiumSource: MARTYRED_BLOODLINE_COMPENDIUM_SOURCE,
      englishName: MARTYRED_BLOODLINE_ENGLISH_NAME
    }),
    findFeat(actor, {
      compendiumSource: RETRIBUTION_COMPENDIUM_SOURCE,
      englishName: RETRIBUTION_ENGLISH_NAME
    }),
    getTimelessSoulFeatureSources(actor, spellItem)
  ]);

  if (state.canUsePassive && state.hasArcaneBloodline) {
    sources.push({
      id: "arcaneBloodline",
      label: state.arcaneBloodlineItem?.name ?? ARCANE_BLOODLINE_ENGLISH_NAME,
      itemUuid: state.arcaneBloodlineItem?.uuid ?? null,
      limited: false,
      persistent: true,
      defaultEnabled: true,
      hasUsesData: true,
      hasRemaining: true
    });
  }

  if (state.metamagicAdeptAvailable) {
    sources.push({
      id: "metamagicAdept",
      label: state.metamagicAdeptItem?.name ?? METAMAGIC_ADEPT_ENGLISH_NAME,
      itemUuid: state.metamagicAdeptItem?.uuid ?? null,
      limited: true,
      persistent: false,
      defaultEnabled: false,
      hasUsesData: state.metamagicAdeptUse.hasUsesData,
      hasRemaining: state.metamagicAdeptUse.hasRemaining,
      usesValue: state.metamagicAdeptUse.remaining,
      usesMax: state.metamagicAdeptUse.max
    });
  }

  if (state.arcaneApotheosisAvailable) {
    sources.push({
      id: "arcaneApotheosis",
      label: state.arcaneApotheosisItem?.name ?? ARCANE_APOTHEOSIS_ENGLISH_NAME,
      itemUuid: state.arcaneApotheosisItem?.uuid ?? null,
      limited: false,
      persistent: true,
      defaultEnabled: false,
      hasUsesData: true,
      hasRemaining: true
    });
  }

  if (state.grandMaestroAvailable) {
    sources.push({
      id: "grandMaestro",
      label: state.grandMaestroItem?.name ?? GRAND_MAESTRO_ENGLISH_NAME,
      itemUuid: state.grandMaestroItem?.uuid ?? null,
      limited: false,
      persistent: true,
      defaultEnabled: true,
      hasUsesData: true,
      hasRemaining: true
    });
  }
  if (state.isSorcererSpell && state.sorcererLevel > 0 && martyredBloodlineItem && retributionItem) {
    sources.push({
      id: "retribution",
      label: retributionItem?.name ?? RETRIBUTION_ENGLISH_NAME,
      itemUuid: retributionItem?.uuid ?? null,
      limited: false,
      persistent: false,
      defaultEnabled: false,
      hasUsesData: true,
      hasRemaining: true,
      dependencies: {
        martyredBloodlineItemUuid: martyredBloodlineItem?.uuid ?? null
      }
    });
  }
  if (Array.isArray(timelessSoulSources) && timelessSoulSources.length) {
    sources.push(...timelessSoulSources);
  }

  return sources;
}

export async function getSorcererArcaneFeatureItem(actor, featureId) {
  const state = await getSorcererArcaneMetamagicState(actor, null);
  if (featureId === "metamagicAdept") return state.metamagicAdeptItem ?? null;
  if (featureId === "arcaneApotheosis") return state.arcaneApotheosisItem ?? null;
  if (featureId === "grandMaestro") return state.grandMaestroItem ?? null;
  return null;
}

export async function getNaniteBloodlineArcanaFeatureSources(actor, spellItem, { durationOverride = null } = {}) {
  if (!actor || !spellItem || spellItem.type !== "spell") return [];
  if (!isSorcererSpellItem(actor, spellItem)) return [];
  const durationEligible = isNaniteBloodlineEligibleSpell(spellItem, { durationOverride });
  if (!durationEligible) return [];

  const naniteBloodlineItem = await findFeat(actor, {
    compendiumSource: NANITE_BLOODLINE_COMPENDIUM_SOURCE,
    englishName: NANITE_BLOODLINE_ENGLISH_NAME
  });
  if (!naniteBloodlineItem) return [];

  return [{
    id: "naniteBloodlineArcana",
    label: naniteBloodlineItem?.name ?? NANITE_BLOODLINE_ENGLISH_NAME,
    itemUuid: naniteBloodlineItem?.uuid ?? null,
    limited: false,
    persistent: true,
    defaultEnabled: true,
    hasUsesData: true,
    hasRemaining: true
  }];
}

export async function getOneBodyTwoMindsFeatureSources(actor, spellItem) {
  if (!actor || !spellItem || spellItem.type !== "spell") return [];
  if (!isSorcererSpellItem(actor, spellItem)) return [];

  const sorcererLevel = getSorcererClassLevel(actor);
  if (sorcererLevel < ONE_BODY_TWO_MINDS_REQUIRED_LEVEL) return [];

  const [possessedBloodlineItem, oneBodyTwoMindsItem] = await Promise.all([
    findFeat(actor, {
      compendiumSource: POSSESSED_BLOODLINE_COMPENDIUM_SOURCE,
      englishName: POSSESSED_BLOODLINE_ENGLISH_NAME
    }),
    findFeat(actor, {
      compendiumSource: ONE_BODY_TWO_MINDS_COMPENDIUM_SOURCE,
      englishName: ONE_BODY_TWO_MINDS_ENGLISH_NAME
    })
  ]);

  if (!possessedBloodlineItem || !oneBodyTwoMindsItem) return [];
  const useState = getUseState(oneBodyTwoMindsItem, { limited: true });

  return [{
    id: "oneBodyTwoMinds",
    label: oneBodyTwoMindsItem?.name ?? ONE_BODY_TWO_MINDS_ENGLISH_NAME,
    itemUuid: oneBodyTwoMindsItem?.uuid ?? null,
    limited: true,
    persistent: false,
    defaultEnabled: false,
    hasUsesData: useState.hasUsesData,
    hasRemaining: useState.hasRemaining,
    usesValue: useState.remaining,
    usesMax: useState.max,
    dependencies: {
      possessedBloodlineItemUuid: possessedBloodlineItem?.uuid ?? null
    }
  }];
}

function getPeerlessSpeedMaxSpellLevel(sorcererLevel) {
  const level = Number(sorcererLevel ?? 0);
  if (!Number.isFinite(level) || level < PEERLESS_SPEED_REQUIRED_LEVEL) return 0;
  const scaled = Math.floor((level - PEERLESS_SPEED_REQUIRED_LEVEL) / 2);
  return Math.max(0, Math.min(8, scaled));
}

export async function getPeerlessSpeedFeatureSources(actor, spellItem) {
  if (!actor || !spellItem || spellItem.type !== "spell") return [];
  if (!isSorcererSpellItem(actor, spellItem)) return [];

  const sorcererLevel = getSorcererClassLevel(actor);
  if (sorcererLevel < PEERLESS_SPEED_REQUIRED_LEVEL) return [];

  const spellLevel = Number(spellItem?.system?.level ?? 0);
  if (!Number.isFinite(spellLevel)) return [];
  const maxEligibleSpellLevel = getPeerlessSpeedMaxSpellLevel(sorcererLevel);
  if (spellLevel > maxEligibleSpellLevel) return [];

  const [astralBloodlineItem, peerlessSpeedItem] = await Promise.all([
    findFeat(actor, {
      compendiumSource: ASTRAL_BLOODLINE_COMPENDIUM_SOURCE,
      englishName: ASTRAL_BLOODLINE_ENGLISH_NAME
    }),
    findFeat(actor, {
      compendiumSource: PEERLESS_SPEED_COMPENDIUM_SOURCE,
      englishName: PEERLESS_SPEED_ENGLISH_NAME
    })
  ]);
  if (!astralBloodlineItem || !peerlessSpeedItem) return [];

  const useState = getUseState(peerlessSpeedItem, { limited: true });
  if (!useState.hasUsesData || !useState.hasRemaining) return [];

  return [{
    id: "peerlessSpeed",
    label: peerlessSpeedItem?.name ?? PEERLESS_SPEED_ENGLISH_NAME,
    itemUuid: peerlessSpeedItem?.uuid ?? null,
    limited: true,
    persistent: false,
    defaultEnabled: false,
    hasUsesData: useState.hasUsesData,
    hasRemaining: useState.hasRemaining,
    usesValue: useState.remaining,
    usesMax: useState.max,
    dependencies: {
      astralBloodlineItemUuid: astralBloodlineItem?.uuid ?? null
    }
  }];
}

export async function getTimelessSoulFeatureSources(actor, spellItem) {
  if (!actor || !spellItem || spellItem.type !== "spell") return [];
  if (!isSorcererSpellItem(actor, spellItem)) return [];

  const sorcererLevel = getSorcererClassLevel(actor);
  if (sorcererLevel < 20) return [];

  const [astralBloodlineItem, timelessSoulItem] = await Promise.all([
    findFeat(actor, {
      compendiumSource: ASTRAL_BLOODLINE_COMPENDIUM_SOURCE,
      englishName: ASTRAL_BLOODLINE_ENGLISH_NAME
    }),
    findFeat(actor, {
      compendiumSource: TIMELESS_SOUL_COMPENDIUM_SOURCE,
      englishName: TIMELESS_SOUL_ENGLISH_NAME
    })
  ]);
  if (!astralBloodlineItem || !timelessSoulItem) return [];

  return [{
    id: "timelessSoul",
    label: timelessSoulItem?.name ?? TIMELESS_SOUL_ENGLISH_NAME,
    itemUuid: timelessSoulItem?.uuid ?? null,
    limited: false,
    persistent: false,
    defaultEnabled: false,
    hasUsesData: true,
    hasRemaining: true,
    requiredMetamagicNames: ["Quicken Spell"],
    autoCheckOnRequiredMetamagicToggle: true,
    dependencies: {
      astralBloodlineItemUuid: astralBloodlineItem?.uuid ?? null
    }
  }];
}
