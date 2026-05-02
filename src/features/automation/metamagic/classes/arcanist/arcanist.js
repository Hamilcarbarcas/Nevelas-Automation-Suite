import { resolveEnglishName } from "../../../utils/compendiumNameResolver.js";
import {
  ARCANE_RESERVOIR_ARCANIST_COMPENDIUM_SOURCE,
  ARCANE_RESERVOIR_EXPLOITER_COMPENDIUM_SOURCE,
  getWizardClassLevel,
  isWizardSpellItem
} from "../wizard/wizard.js";

const METAMIXING_COMPENDIUM_SOURCE = "Compendium.pf1.class-abilities.Item.kUjXL9ec9TGCoJYq";
const METAMIXING_ENGLISH_NAME = "Metamixing";

export const METAMIXING_FEATURE_ID = "metamixing";

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

function getArcanistClassKeys(actor) {
  const classes = actor?.classes ?? {};
  const keys = new Set();
  for (const [classKey, classData] of Object.entries(classes)) {
    const candidates = [classKey, classData?.tag, classData?.name, classData?._id, classData?.id]
      .map(normalizeKey)
      .filter(Boolean);
    if (candidates.some((v) => v.includes("arcanist") || v === "arcn")) {
      keys.add(normalizeKey(classKey));
      candidates.forEach((v) => keys.add(v));
    }
  }
  if (!keys.size) {
    keys.add("arcanist");
    keys.add("arcn");
  }
  return keys;
}

export function getArcanistClassLevel(actor) {
  const classes = actor?.classes ?? {};
  const keys = getArcanistClassKeys(actor);
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

export function isArcanistSpellItem(actor, spellItem) {
  if (!actor || !spellItem || spellItem.type !== "spell") return false;
  const bookClass = getSpellbookClassKey(spellItem);
  if (!bookClass) return false;
  const keys = getArcanistClassKeys(actor);
  if (keys.has(bookClass)) return true;
  return bookClass.includes("arcn");
}

function collectArcaneReservoirItems(actor) {
  const arcanistPool = [];
  const exploiterPool = [];
  for (const item of actor?.items ?? []) {
    if (!isFeatItem(item)) continue;
    const src = (item?._stats?.compendiumSource ?? "").toString();
    if (src === ARCANE_RESERVOIR_ARCANIST_COMPENDIUM_SOURCE) arcanistPool.push(item);
    else if (src === ARCANE_RESERVOIR_EXPLOITER_COMPENDIUM_SOURCE) exploiterPool.push(item);
  }
  return { arcanistPool, exploiterPool };
}

/**
 * Picks the arcane reservoir item whose points Metamixing should spend (arcanist vs exploiter wizard).
 */
export function resolveArcaneReservoirItemForSpell(actor, spellItem) {
  const { arcanistPool, exploiterPool } = collectArcaneReservoirItems(actor);
  const arcanistItem = arcanistPool[0] ?? null;
  const exploiterItem = exploiterPool[0] ?? null;
  if (!arcanistItem && !exploiterItem) return null;
  if (arcanistItem && !exploiterItem) return arcanistItem;
  if (!arcanistItem && exploiterItem) return exploiterItem;
  const arcanistSpell = spellItem ? isArcanistSpellItem(actor, spellItem) : false;
  const wizardSpell = spellItem ? isWizardSpellItem(actor, spellItem) : false;
  if (arcanistSpell && !wizardSpell) return arcanistItem;
  if (wizardSpell && !arcanistSpell) return exploiterItem;
  if (arcanistSpell && wizardSpell) return arcanistItem;
  return arcanistItem;
}

function getReservoirUseSnapshot(item) {
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

export async function getMetamixingState(actor, spellItem = null) {
  if (!actor) {
    return {
      eligible: false,
      metamixingItem: null,
      reservoirItem: null,
      hasUsesData: false,
      usesRemaining: 0,
      usesMax: 0,
      hasRemainingUses: false
    };
  }

  const metamixingItem = await findActorFeat(actor, {
    compendiumSource: METAMIXING_COMPENDIUM_SOURCE,
    englishName: METAMIXING_ENGLISH_NAME
  });

  const reservoirItem = resolveArcaneReservoirItemForSpell(actor, spellItem);
  const arcanistLevel = getArcanistClassLevel(actor);
  const wizardLevel = getWizardClassLevel(actor);
  const spellFromArcanistBook = spellItem ? isArcanistSpellItem(actor, spellItem) : false;
  const spellFromWizardBook = spellItem ? isWizardSpellItem(actor, spellItem) : false;
  const resSrc = (reservoirItem?._stats?.compendiumSource ?? "").toString();

  const arcanistPath =
    Boolean(metamixingItem)
    && Boolean(reservoirItem)
    && arcanistLevel > 0
    && spellFromArcanistBook
    && resSrc === ARCANE_RESERVOIR_ARCANIST_COMPENDIUM_SOURCE;

  const exploiterPath =
    Boolean(metamixingItem)
    && Boolean(reservoirItem)
    && wizardLevel > 0
    && spellFromWizardBook
    && resSrc === ARCANE_RESERVOIR_EXPLOITER_COMPENDIUM_SOURCE;

  const eligible = arcanistPath || exploiterPath;

  const snap = getReservoirUseSnapshot(reservoirItem);

  return {
    eligible,
    metamixingItem,
    reservoirItem,
    hasUsesData: snap.hasUsesData,
    usesRemaining: snap.remaining,
    usesMax: snap.max,
    hasRemainingUses: snap.hasUsesData && snap.remaining > 0
  };
}

export async function getArcanistMetamagicFeatureSources(actor, spellItem) {
  const state = await getMetamixingState(actor, spellItem);
  if (!state.eligible || !state.metamixingItem || !state.reservoirItem) return [];

  return [
    {
      id: METAMIXING_FEATURE_ID,
      label: state.metamixingItem?.name ?? METAMIXING_ENGLISH_NAME,
      itemUuid: state.metamixingItem?.uuid ?? null,
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
