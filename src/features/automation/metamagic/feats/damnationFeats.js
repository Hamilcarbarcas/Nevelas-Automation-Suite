import { MODULE } from "../../../../common/module.js";
import { resolveEnglishName } from "../../utils/compendiumNameResolver.js";
import { contextHasSavingThrow, resolveFeatSaveDcBase } from "../../utils/saveDcUtils.js";

const MALEFICIUM_ID = "maleficium";
const MALEFICIUM_FLAG_KEY = "feats.maleficium";
const MALEFICIUM_COMPENDIUM_SOURCE = "Compendium.pf-content.pf-feats.Item.3eriqSRryCPoqrgI";
const MALEFICIUM_ENGLISH_NAME = "Maleficium";

const DAMNATION_FEAT_UUIDS = new Set([
  MALEFICIUM_COMPENDIUM_SOURCE,
  "Compendium.pf-content.pf-feats.Item.6IKWLLSOXeNyJHjq", // Fiendskin
  "Compendium.pf-content.pf-feats.Item.NPtaLuLvsHp3Tbas", // Mask of Virtue
  "Compendium.pf-content.pf-feats.Item.ppandkNX4ZLUju6l", // Soulless Gaze
]);

function normalizeKey(value) {
  return (value ?? "").toString().trim().toLowerCase();
}

function isEnglishLanguage() {
  return (game?.i18n?.lang ?? "en").toLowerCase().startsWith("en");
}

function isFeatItem(item) {
  const subType = item?.subType ?? item?.system?.subType;
  return item?.type === "feat" && (subType === "feat" || subType === "trait" || subType === "classFeat");
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

function countDamnationFeats(actor) {
  if (!actor?.items) return 0;
  return Array.from(actor.items)
    .filter(isFeatItem)
    .filter((item) => DAMNATION_FEAT_UUIDS.has(item?._stats?.compendiumSource ?? ""))
    .length;
}

/**
 * PF1 stores spell descriptors in system.descriptors (names, base, total); alignment flags
 * on spells are separate. Maleficium keys off the evil descriptor, not system.alignments.
 */
export function spellItemHasEvilDescriptor(item) {
  if (item?.type !== "spell") return false;
  if (item?.system?.alignments?.evil === true) return true;

  const legacy = item?.system?.descriptor;
  if (typeof legacy === "string" && legacySplitHasEvilToken(legacy)) return true;

  const desc = item?.system?.descriptors;
  if (!desc || typeof desc !== "object") return false;

  for (const entry of collectDescriptorLabelStrings(desc)) {
    if (legacySplitHasEvilToken(entry)) return true;
  }

  const total = desc.total;
  if (total && typeof total === "object") {
    if (total instanceof Set) {
      for (const key of total) {
        if (normalizeKey(key) === "evil") return true;
      }
    } else if (!Array.isArray(total)) {
      for (const key of Object.keys(total)) {
        if (total[key] && normalizeKey(key) === "evil") return true;
      }
    }
  }

  return false;
}

function legacySplitHasEvilToken(raw) {
  const s = (raw ?? "").toString().toLowerCase();
  return s.split(",").some((part) => normalizeKey(part) === "evil");
}

function collectDescriptorLabelStrings(descriptors) {
  const out = [];
  if (Array.isArray(descriptors.names)) out.push(...descriptors.names);
  if (Array.isArray(descriptors.base)) out.push(...descriptors.base);
  const value = descriptors.value;
  if (typeof value === "string" && value.trim()) out.push(value);
  return out;
}

function isEvilDescriptorSpell(context, action) {
  const item = action?.item;
  if (item?.type === "spell" && spellItemHasEvilDescriptor(item)) {
    return true;
  }

  const contextEvil = context?.alignments?.resolved?.evil;
  if (typeof contextEvil === "boolean") {
    return contextEvil;
  }
  const actionEvil = action?.action?.alignments?.evil;
  if (typeof actionEvil === "boolean") {
    return actionEvil;
  }
  return spellItemHasEvilDescriptor(item) || Boolean(item?.system?.alignments?.evil);
}

function getDcBonusByCount(damnationCount) {
  if (damnationCount >= 3) return 2;
  if (damnationCount >= 1) return 1;
  return 0;
}

export function getMaleficiumOptionId() {
  return MALEFICIUM_ID;
}

export function getMaleficiumState(actor) {
  return actor?.getFlag?.(MODULE.ID, MALEFICIUM_FLAG_KEY) === true;
}

export async function setMaleficiumState(actor, enabled) {
  if (!actor || typeof actor?.setFlag !== "function") return false;
  const nextValue = enabled === true;
  await actor.setFlag(MODULE.ID, MALEFICIUM_FLAG_KEY, nextValue);
  return nextValue;
}

export async function getMaleficiumSource(actor) {
  if (!actor) return null;
  const featItem = await findActorFeat(actor, {
    compendiumSource: MALEFICIUM_COMPENDIUM_SOURCE,
    englishName: MALEFICIUM_ENGLISH_NAME
  });
  if (!featItem) return null;
  return {
    id: MALEFICIUM_ID,
    label: featItem?.name ?? MALEFICIUM_ENGLISH_NAME,
    itemUuid: featItem?.uuid ?? null,
    persistent: true,
    damnationCount: countDamnationFeats(actor)
  };
}

export async function prepareMaleficiumContext(action, context) {
  const actor = context?.actor ?? action?.actor ?? action?.token?.actor ?? null;
  if (!actor || !action?.item) return;
  if (action.item.type !== "spell") return;
  const evilEligible = isEvilDescriptorSpell(context, action);
  if (!evilEligible) return;

  const source = await getMaleficiumSource(actor);
  if (!source) return;

  const formEnabled = context?.featOptions?.[MALEFICIUM_ID]?.enabled;
  const enabled = formEnabled === true || (formEnabled === undefined && getMaleficiumState(actor) === true);
  if (!enabled) return;

  const damnationCount = Number(source?.damnationCount ?? countDamnationFeats(actor));
  if (!Number.isFinite(damnationCount) || damnationCount <= 0) return;

  const baseSpellLevel = Number(context?.spellLevel?.original ?? action?.item?.system?.level ?? 0);
  const dcBonus = getDcBonusByCount(damnationCount);
  const metamagicSlotAdjustment = damnationCount >= 2 ? -1 : 0;
  const casterLevelBonus = damnationCount >= 4 ? 2 : 0;
  const minimumConsumedSlotLevel = damnationCount >= 2 && Number.isFinite(baseSpellLevel)
    ? baseSpellLevel + 1
    : null;

  context.featEffects ??= {};
  context.featEffects[MALEFICIUM_ID] = {
    active: true,
    label: source.label,
    damnationCount,
    dcBonus,
    metamagicSlotAdjustment,
    casterLevelBonus,
    minimumConsumedSlotLevel
  };

  if (casterLevelBonus > 0) {
    const currentCl = Number(action?.shared?.rollData?.cl ?? 0);
    if (Number.isFinite(currentCl)) {
      action.shared.rollData.cl = currentCl + casterLevelBonus;
    }
  }
}

export function getMaleficiumSlotAdjustment(context, { hasAppliedMetamagic = false } = {}) {
  if (!hasAppliedMetamagic) return 0;
  const adjustment = Number(context?.featEffects?.[MALEFICIUM_ID]?.metamagicSlotAdjustment ?? 0);
  if (!Number.isFinite(adjustment)) return 0;
  return adjustment;
}

export function getMaleficiumMinimumConsumedSlotLevel(context, { hasAppliedMetamagic = false } = {}) {
  if (!hasAppliedMetamagic) return null;
  const value = Number(context?.featEffects?.[MALEFICIUM_ID]?.minimumConsumedSlotLevel);
  if (!Number.isFinite(value)) return null;
  return value;
}

export function applyMaleficiumPostMetamagic(action, context) {
  const effect = context?.featEffects?.[MALEFICIUM_ID];
  if (!effect?.active) return;
  const dcBonus = Number(effect?.dcBonus ?? 0);
  if (!Number.isFinite(dcBonus) || dcBonus <= 0) return;
  if (!contextHasSavingThrow(context, action)) return;

  const currentDc = resolveFeatSaveDcBase(action, context);
  if (!Number.isFinite(currentDc)) return;
  const nextDc = currentDc + dcBonus;

  context.save ??= {};
  const existingBase = Number(context.save.baseDc);
  if (!Number.isFinite(existingBase)) {
    context.save.baseDc = currentDc;
  }
  context.save.dc = nextDc;
  const evaluatedTotal = Number(context?.save?.evaluated?.total);
  if (Number.isFinite(evaluatedTotal)) {
    context.save.evaluated.total = evaluatedTotal + dcBonus;
  } else {
    context.save.evaluated ??= {};
    context.save.evaluated.total = nextDc;
  }

  action.shared ??= {};
  action.shared.saveDC = nextDc;
}
