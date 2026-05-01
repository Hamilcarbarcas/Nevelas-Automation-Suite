import { resolveEnglishName } from "../../utils/compendiumNameResolver.js";
import { METAMAGIC_DEFINITION as ExtendSpell } from "../extendSpell.js";
import { resolveMetamagicNameFromDatabase } from "../metamagic.js";

export const MASK_FOCUS_ID = "maskFocus";
export const MASK_FOCUS_FEATURE_ID = "maskFocus";
const MASK_FOCUS_COMPENDIUM_SOURCE = "Compendium.pf-content.pf-feats.Item.31VNHRs8WTUIAcEU";
const MASK_FOCUS_ENGLISH_NAME = "Mask Focus";

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

function validateMaskFocusUses(featItem) {
  const uses = featItem?.system?.uses;
  if (!uses || !uses.per) {
    return { ok: false, reason: "missingUsesData" };
  }
  const remaining = Number(uses.value ?? 0);
  if (!Number.isFinite(remaining) || remaining <= 0) {
    return { ok: false, reason: "noRemainingUses" };
  }
  return { ok: true };
}

function spellbookIsSpellLike(spellItem) {
  return normalizeKey(spellItem?.system?.spellbook) === "spelllike";
}

export async function getMaskFocusSource(actor) {
  if (!actor) return null;
  const featItem = await findActorFeat(actor, {
    compendiumSource: MASK_FOCUS_COMPENDIUM_SOURCE,
    englishName: MASK_FOCUS_ENGLISH_NAME
  });
  if (!featItem) return null;
  const uses = featItem?.system?.uses;
  if (!uses?.per) return null;
  const usesMax = Number(uses.max ?? 0);
  if (!Number.isFinite(usesMax) || usesMax < 1) return null;
  const usesValue = Number(uses?.value ?? 0);
  return {
    id: MASK_FOCUS_ID,
    label: featItem?.name ?? MASK_FOCUS_ENGLISH_NAME,
    itemUuid: featItem?.uuid ?? null,
    limited: true,
    usesValue: Number.isFinite(usesValue) ? usesValue : 0,
    usesMax: Number.isFinite(usesMax) ? usesMax : 0,
    hasUsesData: Boolean(uses?.per),
    hasRemaining: Number.isFinite(usesValue) && usesValue > 0
  };
}

export async function prepareMaskFocusContext(action, context) {
  const actor = context?.actor ?? action?.actor ?? action?.token?.actor ?? null;
  if (!actor || !action?.item) return;
  if (action.item.type !== "spell") return;

  const formEnabled = context?.featOptions?.[MASK_FOCUS_ID]?.enabled === true;
  if (!formEnabled) return;

  const names = Array.isArray(context?.metamagicNames) ? context.metamagicNames : [];
  const normalizedNames = names.map((n) => resolveMetamagicNameFromDatabase(n) ?? n);
  if (!normalizedNames.includes(ExtendSpell.name)) {
    ui.notifications.warn(game.i18n.localize("NAS.metamagic.maskFocus.requiresExtendSpell"));
    return;
  }

  if (spellbookIsSpellLike(action.item)) {
    ui.notifications.warn(game.i18n.localize("NAS.metamagic.maskFocus.notOnSpellLike"));
    return;
  }

  const featItem = await findActorFeat(actor, {
    compendiumSource: MASK_FOCUS_COMPENDIUM_SOURCE,
    englishName: MASK_FOCUS_ENGLISH_NAME
  });
  if (!featItem) {
    ui.notifications.warn(game.i18n.localize("NAS.metamagic.maskFocus.featNotFound"));
    return;
  }

  const useValidation = validateMaskFocusUses(featItem);
  if (!useValidation.ok) {
    if (useValidation.reason === "missingUsesData") {
      ui.notifications.warn(game.i18n.format("NAS.metamagic.maskFocus.missingUses", { name: featItem.name }));
    } else {
      ui.notifications.warn(game.i18n.localize("NAS.metamagic.maskFocus.noUsesRemaining"));
    }
    return;
  }

  context.featEffects ??= {};
  context.featEffects[MASK_FOCUS_ID] = {
    active: true,
    label: featItem?.name ?? MASK_FOCUS_ENGLISH_NAME,
    itemUuid: featItem?.uuid ?? null
  };
}
