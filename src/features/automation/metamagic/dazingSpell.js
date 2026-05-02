import { resolveEnglishName } from "../utils/compendiumNameResolver.js";
import { getAlignmentValue, getCreatureTypeValues } from "../utils/creatureTypeUtils.js";

export const METAMAGIC_DEFINITION = {
  key: "dazingSpell",
  name: "Dazing Spell",
  get prefix() {
    return globalThis.game?.i18n?.localize?.("NAS.metamagic.prefixes.dazingSpell") ?? "Dazing";
  },
};

export const DAZING_FEAT_ROUNDS = 1;
const DAZE_SPELL_NAMES = new Set([
  "Arrow of Law",
  "Ear-Piercing Scream",
  "Order's Wrath",
  "Sheet Lightning",
  "Telekinetic Storm",
]);

function isEnglishLanguage() {
  return (game?.i18n?.lang ?? "en").toLowerCase().startsWith("en");
}

function isChaoticAlignment(alignment) {
  const value = getAlignmentValue(alignment);
  return value.includes("c");
}

function isOutsider(actor) {
  const values = getCreatureTypeValues(actor);
  return values.some((value) => value.includes("outsider"));
}

export async function getSpellEnglishName(item) {
  const itemName = (item?.name ?? "").toString().trim();
  if (!itemName) return itemName;
  if (isEnglishLanguage()) return itemName;
  const canUseBabele = game?.modules?.get("babele")?.active;
  if (canUseBabele && typeof item?.getFlag === "function") {
    const originalName = item.getFlag("babele", "originalName");
    if (originalName) return originalName;
  }
  return resolveEnglishName(itemName, { documentName: "Item", deepScanMode: "off" });
}

export function getDazingExtraRoundsForTarget(spellName, targetActor) {
  if (!spellName || !DAZE_SPELL_NAMES.has(spellName)) return 0;
  if (spellName === "Arrow of Law") {
    if (!targetActor) return 0;
    const alignment = targetActor?.system?.details?.alignment;
    if (!isChaoticAlignment(alignment)) return 0;
    if (!isOutsider(targetActor)) return 0;
  }
  return 1;
}

export async function applyDazingSpell(context, action, options = {}) {
  if (!context) return false;
  const hasDamage = context?.damage?.hasDamage === true;
  const parts = context?.damage?.parts ?? [];
  const hasFormula =
    Array.isArray(parts) &&
    parts.some((part) => typeof (part?.formula ?? part?.[0]) === "string");
  if (!hasDamage && !hasFormula) return false;

  if (!context.metamagic) {
    context.metamagic = { applied: [], slotIncrease: 0 };
  }

  if (!context.metamagic.applied.includes(METAMAGIC_DEFINITION.name)) {
    context.metamagic.applied.push(METAMAGIC_DEFINITION.name);
    context.metamagic.slotIncrease += 3;
  }

  context.metamagic.dazing = true;
  const roundsRaw = options?.rounds;
  const roundsValue = Number.isFinite(roundsRaw) ? Math.max(0, Math.floor(roundsRaw)) : DAZING_FEAT_ROUNDS;
  context.metamagic.dazingRounds = roundsValue;

  if (!context.save?.type) {
    context.save ??= {};
    context.save.type = "will";
    context.save.description = context.save.description || "Will negates";
  }

  const item = action?.item ?? null;
  if (item) {
    context.metamagic.dazingSpellName = await getSpellEnglishName(item);
  }

  return true;
}
