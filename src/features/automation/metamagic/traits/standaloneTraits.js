import { resolveEnglishName } from "../../utils/compendiumNameResolver.js";
import {
  findChoiceByIdentifier,
  findChoiceByStoredName,
  getDictionaryString,
  setDictionaryEntries
} from "../../utils/itemDictionarySelection.js";
import { getRacialSpellLikeTraitSources } from "./racialSpellLikeTraits.js";

const MAGICAL_LINEAGE_COMPENDIUM_SOURCE = "Compendium.pf-content.pf-traits.Item.kDx6ZGhbIteoi7lG";
const MAGICAL_LINEAGE_ENGLISH_NAME = "Magical Lineage";
const MAGICAL_LINEAGE_ID = "magicalLineage";
const MAGICAL_LINEAGE_SELECTED_SPELL_NAME_FLAG = "MLSSN";
const TRANSMUTER_OF_KORADA_COMPENDIUM_SOURCE = "Compendium.pf-content.pf-traits.Item.3wlB6o57m3XQX5fu";
const TRANSMUTER_OF_KORADA_ENGLISH_NAME = "Transmuter of Korada";
export const TRANSMUTER_OF_KORADA_ID = "transmuterOfKorada";
const TRANSMUTER_OF_KORADA_SELECTED_SPELL_FLAG = "ToKSelectedSpell";
const TRANSMUTATION_SCHOOL_KEYS = new Set(["trs", "transmutation", "tra"]);
const TRANSMUTER_OF_KORADA_SPELL_UUIDS = [
  "Compendium.pf1.spells.Item.usdv1eqvibmxun6x",
  "Compendium.pf1.spells.Item.05i5rxwim12hwktu",
  "Compendium.pf1.spells.Item.ns8jbp0ilbvbvuif",
  "Compendium.pf1.spells.Item.d4oubr5bdoo8w1ev",
  "Compendium.pf1.spells.Item.743anqr1ahefv8zd",
  "Compendium.pf1.spells.Item.b9ggsagifzk4fwut"
];

let transmuterOfKoradaHookRegistered = false;
let magicalLineageHookRegistered = false;

function localizeMetamagic(path) {
  return game.i18n.localize(`NAS.metamagic.${path}`);
}

function formatMetamagic(path, data = {}) {
  return game.i18n.format(`NAS.metamagic.${path}`, data);
}

function normalizeKey(value) {
  return (value ?? "").toString().trim().toLowerCase();
}

function isEnglishLanguage() {
  return (game?.i18n?.lang ?? "en").toLowerCase().startsWith("en");
}

function isFeatLikeItem(item) {
  const subType = item?.subType ?? item?.system?.subType;
  return item?.type === "feat" && (subType === "trait" || subType === "feat" || subType === "classFeat");
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

function getFastMatchReason(item, { compendiumSource, englishName }) {
  const source = item?._stats?.compendiumSource ?? "";
  if (source && source === compendiumSource) return "compendiumSource";
  const originalName = getBabeleOriginalName(item);
  if (originalName && originalName === englishName) return "babele.originalName";
  if (isEnglishLanguage() && (item?.name ?? "") === englishName) return "name";
  return null;
}

async function findActorTrait(actor, { compendiumSource, englishName }) {
  const items = actor?.items;
  if (!items) return null;
  const feats = Array.from(items).filter(isFeatLikeItem);

  for (const feat of feats) {
    const reason = getFastMatchReason(feat, { compendiumSource, englishName });
    if (reason) return feat;
  }

  if (isEnglishLanguage()) return null;
  for (const feat of feats) {
    const resolved = await resolveEnglishName(feat?.name, { documentName: "Item", deepScanMode: "off" });
    if ((resolved ?? "") === englishName) return feat;
  }
  return null;
}

function extractDocumentIdFromUuid(uuid) {
  const value = (uuid ?? "").toString().trim();
  if (!value) return "";
  const parts = value.split(".");
  return parts.length ? (parts[parts.length - 1] ?? "") : "";
}

function getSpellMatchKeys(spellItem) {
  if (!spellItem) return new Set();
  const keys = new Set();
  const uuid = (spellItem.uuid ?? "").toString().trim();
  const id = (spellItem.id ?? "").toString().trim();
  if (uuid) keys.add(uuid);
  if (id) keys.add(id);
  return keys;
}

function selectedSpellMatchesAction(actionItem, selectedSpellUuid) {
  if (!actionItem || !selectedSpellUuid) return false;
  const selectedId = extractDocumentIdFromUuid(selectedSpellUuid);
  const spellKeys = getSpellMatchKeys(actionItem);
  if (spellKeys.has(selectedSpellUuid)) return true;
  if (selectedId && spellKeys.has(selectedId)) return true;
  return false;
}

function getActorSpellChoices(actor) {
  if (!actor?.items) return [];
  return Array.from(actor.items)
    .filter((item) => item?.type === "spell")
    .filter((item) => !isSpellLikeSpellbook(item))
    .map((item) => ({
      uuid: item?.uuid ?? "",
      id: item?.id ?? "",
      label: item?.name ?? "Spell",
      img: item?.img ?? ""
    }))
    .filter((entry) => entry.uuid || entry.id)
    .sort((a, b) => (a?.label ?? "").localeCompare(b?.label ?? ""));
}

function getUseState(item) {
  const uses = item?.system?.uses;
  if (!uses || !uses.per) {
    return {
      limited: false,
      hasUsesData: false,
      remaining: null,
      max: null,
      hasRemaining: true
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

async function resolveSpellNameCandidates(spellItem) {
  const localized = spellItem?.name ?? "";
  const babeleOriginal = getBabeleOriginalName(spellItem);
  const resolvedEnglish = await resolveEnglishName(localized, { documentName: "Item", deepScanMode: "off" });
  return [localized, babeleOriginal, resolvedEnglish]
    .map(normalizeKey)
    .filter(Boolean);
}

function isCklModuleActive() {
  return game?.modules?.get?.("ckl-roll-bonuses")?.active === true;
}

function spellIsTransmutationSchool(spellItem) {
  const school = normalizeKey(spellItem?.system?.school);
  return TRANSMUTATION_SCHOOL_KEYS.has(school);
}

function getCklTransmuterConfig(item) {
  const moduleFlags = item?.flags?.["ckl-roll-bonuses"] ?? {};
  const booleanFlags = item?.system?.flags?.boolean ?? {};
  const hasBonusBoolean = booleanFlags?.bonus_cl === true;
  const hasSchoolBoolean = booleanFlags?.["target_spell-school"] === true;
  const bonusClFormula = (moduleFlags?.bonus_cl ?? "").toString().trim();
  const targetSchools = Array.isArray(moduleFlags?.["target_spell-school"])
    ? moduleFlags["target_spell-school"]
    : moduleFlags?.["target_spell-school"]
      ? [moduleFlags["target_spell-school"]]
      : [];
  const normalizedTargetSchools = targetSchools.map(normalizeKey).filter(Boolean);
  const hasTransmutationTarget = normalizedTargetSchools.some((value) => TRANSMUTATION_SCHOOL_KEYS.has(value));
  return {
    hasBonusBoolean,
    hasSchoolBoolean,
    bonusClFormula,
    hasTransmutationTarget
  };
}

function hasConfiguredCklTransmuterCasterLevelBonus(item) {
  if (!isCklModuleActive()) return false;
  const cfg = getCklTransmuterConfig(item);
  return cfg.hasBonusBoolean && cfg.hasSchoolBoolean && cfg.bonusClFormula.length > 0 && cfg.hasTransmutationTarget;
}

export function getTransmuterOfKoradaSelectedSpellName(item) {
  return getDictionaryString(item, TRANSMUTER_OF_KORADA_SELECTED_SPELL_FLAG, { normalize: true });
}

export async function setTransmuterOfKoradaSelectedSpellName(item, selectedSpellName) {
  if (!item || typeof item?.update !== "function") return "";
  const normalized = normalizeKey(selectedSpellName);
  await setDictionaryEntries(item, {
    [TRANSMUTER_OF_KORADA_SELECTED_SPELL_FLAG]: normalized
  });
  return normalized;
}

async function getTransmuterOfKoradaSpellChoices() {
  const docs = await Promise.all(
    TRANSMUTER_OF_KORADA_SPELL_UUIDS.map(async (uuid) => {
      try {
        return await fromUuid(uuid);
      } catch (_err) {
        return null;
      }
    })
  );
  return docs
    .map((doc, index) => {
      const uuid = TRANSMUTER_OF_KORADA_SPELL_UUIDS[index] ?? "";
      if (!doc) return null;
      return {
        uuid: doc.uuid ?? uuid,
        id: doc.id ?? "",
        label: (doc.name ?? "").toString().trim(),
        img: doc.img ?? ""
      };
    })
    .filter((entry) => entry?.label && (entry?.uuid || entry?.id))
    .sort((a, b) => a.label.localeCompare(b.label));
}

async function isTransmuterOfKoradaItem(item) {
  if (!isFeatLikeItem(item)) return false;
  const reason = getFastMatchReason(item, {
    compendiumSource: TRANSMUTER_OF_KORADA_COMPENDIUM_SOURCE,
    englishName: TRANSMUTER_OF_KORADA_ENGLISH_NAME
  });
  if (reason) return true;
  if (isEnglishLanguage()) return false;
  const resolved = await resolveEnglishName(item?.name, { documentName: "Item", deepScanMode: "off" });
  return (resolved ?? "") === TRANSMUTER_OF_KORADA_ENGLISH_NAME;
}

async function isMagicalLineageItem(item) {
  if (!isFeatLikeItem(item)) return false;
  const reason = getFastMatchReason(item, {
    compendiumSource: MAGICAL_LINEAGE_COMPENDIUM_SOURCE,
    englishName: MAGICAL_LINEAGE_ENGLISH_NAME
  });
  if (reason) return true;
  if (isEnglishLanguage()) return false;
  const resolved = await resolveEnglishName(item?.name, { documentName: "Item", deepScanMode: "off" });
  return (resolved ?? "") === MAGICAL_LINEAGE_ENGLISH_NAME;
}

function promptMagicalLineageSpellSelection(actor, { currentSpellName = "" } = {}) {
  const choices = getActorSpellChoices(actor);
  if (!choices.length) {
    ui.notifications.warn(localizeMetamagic("warnings.noNonSpelllikeSpellsFound"));
    return Promise.resolve(null);
  }
  const currentChoice = findChoiceByStoredName(choices, currentSpellName);
  const currentValue = currentChoice?.uuid || currentChoice?.id || "";
  const cards = choices
    .map((choice, index) => {
      const value = choice.uuid || choice.id;
      const checked = value === currentValue;
      const img = choice.img ? `<img src="${choice.img}" style="width:32px;height:32px;border-radius:4px;" />` : "";
      const safeLabel = foundry.utils.escapeHTML(choice.label ?? `Spell ${index + 1}`);
      return `
        <label style="display:flex;align-items:center;gap:8px;padding:4px 0;">
          <input type="radio" name="magicalLineageSpell" value="${value}" ${checked ? "checked" : ""} />
          ${img}
          <span>${safeLabel}</span>
        </label>
      `;
    })
    .join("");
  return new Promise((resolve) => {
    new Dialog({
      title: MAGICAL_LINEAGE_ENGLISH_NAME,
      content: `
        <form>
          <p style="margin:0 0 8px 0;">${formatMetamagic("dialogs.chooseSpellForName", {
            name: game.i18n.localize("NAS.metamagic.featureNames.magicalLineage")
          })}</p>
          <div style="max-height:320px;overflow-y:auto;border:1px solid #888;border-radius:4px;padding:6px 8px;">
            ${cards}
          </div>
        </form>
      `,
      buttons: {
        ok: {
          label: game.i18n.localize("NAS.common.buttons.select"),
          callback: (html) => {
            const selected = `${html.find('input[name="magicalLineageSpell"]:checked').val() ?? ""}`.trim();
            if (!selected) {
              resolve(null);
              return;
            }
            const match = findChoiceByIdentifier(choices, selected);
            if (!match?.label) {
              resolve(null);
              return;
            }
            resolve({
              selectedSpellName: match.label,
              selectedSpellLabel: match.label
            });
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

export async function promptTransmuterOfKoradaSpellSelection(item, { currentSpellName = "" } = {}) {
  const choices = await getTransmuterOfKoradaSpellChoices();
  if (!choices.length) {
    ui.notifications.warn(localizeMetamagic("warnings.noValidTransmuterChoices"));
    return null;
  }
  const currentName = currentSpellName || getTransmuterOfKoradaSelectedSpellName(item);
  const currentChoice = findChoiceByStoredName(choices, currentName);
  const currentValue = currentChoice?.uuid || currentChoice?.id || "";
  const cards = choices
    .map((choice, index) => {
      const value = choice.uuid || choice.id;
      const checked = value === currentValue;
      const img = choice.img ? `<img src="${choice.img}" style="width:32px;height:32px;border-radius:4px;" />` : "";
      const safeLabel = foundry.utils.escapeHTML(choice.label ?? `Spell ${index + 1}`);
      return `
        <label class="nas-transmuter-spell-option" style="display:flex;align-items:center;gap:8px;padding:4px 0;">
          <input type="radio" name="transmuterOfKoradaSpell" value="${value}" ${checked ? "checked" : ""} />
          ${img}
          <span>${safeLabel}</span>
        </label>
      `;
    })
    .join("");
  return new Promise((resolve) => {
    new Dialog({
      title: item?.name ?? TRANSMUTER_OF_KORADA_ENGLISH_NAME,
      content: `
        <form>
          <p style="margin:0 0 8px 0;">${formatMetamagic("dialogs.chooseSpellForName", {
            name: foundry.utils.escapeHTML(item?.name ?? TRANSMUTER_OF_KORADA_ENGLISH_NAME)
          })}</p>
          <div style="max-height:320px;overflow-y:auto;border:1px solid #888;border-radius:4px;padding:6px 8px;">
            ${cards}
          </div>
        </form>
      `,
      buttons: {
        ok: {
          label: game.i18n.localize("NAS.common.buttons.select"),
          callback: (html) => {
            const selected = `${html.find('input[name="transmuterOfKoradaSpell"]:checked').val() ?? ""}`.trim();
            if (!selected) {
              resolve(null);
              return;
            }
            const selectedSpell = findChoiceByIdentifier(choices, selected);
            const selectedSpellName = normalizeKey(selectedSpell?.label ?? "");
            if (!selectedSpellName || !selectedSpell?.label) {
              resolve(null);
              return;
            }
            resolve({
              selectedSpellName,
              selectedSpellLabel: selectedSpell.label
            });
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

export async function getTransmuterOfKoradaSource(actor, spellItem) {
  if (!actor || !spellItem || spellItem.type !== "spell") {
    return null;
  }
  const traitItem = await findActorTrait(actor, {
    compendiumSource: TRANSMUTER_OF_KORADA_COMPENDIUM_SOURCE,
    englishName: TRANSMUTER_OF_KORADA_ENGLISH_NAME
  });
  if (!traitItem) {
    return null;
  }
  const use = getUseState(traitItem);
  const selectedSpellName = getTransmuterOfKoradaSelectedSpellName(traitItem);
  const spellCandidates = await resolveSpellNameCandidates(spellItem);
  const selectedSpellMatches = selectedSpellName
    ? spellCandidates.includes(selectedSpellName)
    : false;
  const spellIsTransmutation = spellIsTransmutationSchool(spellItem);
  const cklConfigured = spellIsTransmutation && hasConfiguredCklTransmuterCasterLevelBonus(traitItem);
  return {
    id: TRANSMUTER_OF_KORADA_ID,
    label: traitItem?.name ?? TRANSMUTER_OF_KORADA_ENGLISH_NAME,
    itemUuid: traitItem?.uuid ?? null,
    itemId: traitItem?.id ?? null,
    limited: use.limited,
    hasUsesData: use.hasUsesData,
    hasRemaining: use.hasRemaining,
    usesValue: use.remaining,
    usesMax: use.max,
    effectType: "transmuterOfKorada",
    requiresMetamagicIntent: false,
    blocksExtendLike: true,
    selectedSpellName,
    selectedSpellMatches,
    requiresSpellSelection: !selectedSpellName,
    casterLevelBonus: spellIsTransmutation && !cklConfigured ? 1 : 0
  };
}

export function registerTransmuterOfKoradaItemHook() {
  if (transmuterOfKoradaHookRegistered) return;
  transmuterOfKoradaHookRegistered = true;
  Hooks.on("createItem", async (item, _options, userId) => {
    if (!item || userId !== game?.user?.id) return;
    if (!item?.parent || item.parent.documentName !== "Actor") return;
    if (!item?.isOwner) return;
    if (!(await isTransmuterOfKoradaItem(item))) return;
    if (getTransmuterOfKoradaSelectedSpellName(item)) return;
    const selection = await promptTransmuterOfKoradaSpellSelection(item);
    if (!selection?.selectedSpellName) return;
    await setTransmuterOfKoradaSelectedSpellName(item, selection.selectedSpellName);
  });
}

export function registerMagicalLineageItemHook() {
  if (magicalLineageHookRegistered) return;
  magicalLineageHookRegistered = true;
  Hooks.on("createItem", async (item, _options, userId) => {
    if (!item || userId !== game?.user?.id) return;
    if (!item?.parent || item.parent.documentName !== "Actor") return;
    if (!item?.isOwner) return;
    if (!(await isMagicalLineageItem(item))) return;
    if (getDictionaryString(item, MAGICAL_LINEAGE_SELECTED_SPELL_NAME_FLAG)) return;
    const selection = await promptMagicalLineageSpellSelection(item.parent);
    if (!selection?.selectedSpellName) return;
    await setMagicalLineageState(item.parent, {
      enabled: false,
      selectedSpellName: selection.selectedSpellName,
      selectedSpellLabel: selection.selectedSpellLabel
    });
  });
}

export function getMagicalLineageOptionId() {
  return MAGICAL_LINEAGE_ID;
}

export async function getMagicalLineageSource(actor) {
  if (!actor) return null;
  const traitItem = await findActorTrait(actor, {
    compendiumSource: MAGICAL_LINEAGE_COMPENDIUM_SOURCE,
    englishName: MAGICAL_LINEAGE_ENGLISH_NAME
  });
  if (!traitItem) return null;
  return {
    id: MAGICAL_LINEAGE_ID,
    label: traitItem?.name ?? MAGICAL_LINEAGE_ENGLISH_NAME,
    itemUuid: traitItem?.uuid ?? null,
    persistent: true
  };
}

export async function getMagicalLineageState(actor) {
  const traitItem = await findActorTrait(actor, {
    compendiumSource: MAGICAL_LINEAGE_COMPENDIUM_SOURCE,
    englishName: MAGICAL_LINEAGE_ENGLISH_NAME
  });
  const selectedSpellName = getDictionaryString(traitItem, MAGICAL_LINEAGE_SELECTED_SPELL_NAME_FLAG);
  return {
    enabled: false,
    selectedSpellUuid: "",
    selectedSpellLabel: selectedSpellName,
    selectedSpellName
  };
}

export async function setMagicalLineageState(actor, nextState = {}) {
  const traitItem = await findActorTrait(actor, {
    compendiumSource: MAGICAL_LINEAGE_COMPENDIUM_SOURCE,
    englishName: MAGICAL_LINEAGE_ENGLISH_NAME
  });
  if (!traitItem) {
    return {
      enabled: false,
      selectedSpellUuid: "",
      selectedSpellLabel: "",
      selectedSpellName: ""
    };
  }
  const current = await getMagicalLineageState(actor);
  const nextSelectedSpellName = (
    nextState.selectedSpellName
    ?? nextState.selectedSpellLabel
    ?? current.selectedSpellName
    ?? ""
  ).toString().trim();
  await setDictionaryEntries(traitItem, {
    [MAGICAL_LINEAGE_SELECTED_SPELL_NAME_FLAG]: nextSelectedSpellName
  });
  const merged = {
    enabled: false,
    selectedSpellUuid: "",
    selectedSpellLabel: nextSelectedSpellName,
    selectedSpellName: nextSelectedSpellName
  };
  return merged;
}

export async function prepareMagicalLineageContext(action, context) {
  const actor = context?.actor ?? action?.actor ?? action?.token?.actor ?? null;
  if (!actor || !action?.item) return;
  const source = await getMagicalLineageSource(actor);
  if (!source) return;

  const optionRaw = context?.featOptions?.[MAGICAL_LINEAGE_ID] ?? await getMagicalLineageState(actor);
  const optionState = {
    enabled: optionRaw?.enabled === true,
    selectedSpellUuid: (optionRaw?.selectedSpellUuid ?? "").toString().trim(),
    selectedSpellLabel: (optionRaw?.selectedSpellLabel ?? "").toString().trim(),
    selectedSpellName: (optionRaw?.selectedSpellName ?? optionRaw?.selectedSpellLabel ?? "").toString().trim()
  };
  if (!optionState.enabled) return;
  let selectedSpellMatches = false;
  if (optionState.selectedSpellUuid) {
    selectedSpellMatches = selectedSpellMatchesAction(action.item, optionState.selectedSpellUuid);
  } else if (optionState.selectedSpellName) {
    const spellCandidates = await resolveSpellNameCandidates(action.item);
    selectedSpellMatches = spellCandidates.includes(normalizeKey(optionState.selectedSpellName));
  }
  if (!selectedSpellMatches) return;

  context.featEffects ??= {};
  context.featEffects[MAGICAL_LINEAGE_ID] = {
    active: true,
    label: source.label,
    spellUuid: optionState.selectedSpellUuid || action?.item?.uuid || "",
    spellLabel: optionState.selectedSpellLabel || optionState.selectedSpellName || action?.item?.name || "",
    metamagicSlotAdjustment: -1
  };
}

export function getMagicalLineageSlotAdjustment(context, { hasAppliedMetamagic = false } = {}) {
  if (!hasAppliedMetamagic) return 0;
  const adjustment = Number(
    context?.featEffects?.[MAGICAL_LINEAGE_ID]?.metamagicSlotAdjustment ?? 0
  );
  if (!Number.isFinite(adjustment)) return 0;
  return adjustment;
}

export { getRacialSpellLikeTraitSources };
