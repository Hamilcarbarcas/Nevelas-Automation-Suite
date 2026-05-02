import { resolveEnglishName } from "../../utils/compendiumNameResolver.js";
import {
  findChoiceByStoredName,
  getDictionaryString,
  setDictionaryEntries
} from "../../utils/itemDictionarySelection.js";

const WAYANG_SPELLHUNTER_COMPENDIUM_SOURCE = "Compendium.pf-content.pf-traits.Item.XgYRyJKO8F3XWZO4";
const WAYANG_SPELLHUNTER_ENGLISH_NAME = "Wayang Spellhunter (Minata)";
export const WAYANG_SPELLHUNTER_ID = "wayangSpellhunterMinata";
const WAYANG_SELECTED_SPELL_NAME_FLAG = "WSSelectedSpellName";
const WAYANG_MAX_SPELL_LEVEL = 3;

let wayangSpellhunterHookRegistered = false;

function localizeMetamagic(path) {
  return game.i18n.localize(`NAS.metamagic.${path}`);
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

function isSpellLikeSpellbook(spellItem) {
  return normalizeKey(spellItem?.system?.spellbook) === "spelllike";
}

function getWayangSpellhunterSpellChoices(actor) {
  if (!actor?.items) return [];
  return Array.from(actor.items)
    .filter((item) => item?.type === "spell")
    .filter((item) => !isSpellLikeSpellbook(item))
    .filter((item) => {
      const level = Number(item?.system?.level ?? 0);
      return Number.isFinite(level) && level <= WAYANG_MAX_SPELL_LEVEL;
    })
    .map((item) => ({
      uuid: item?.uuid ?? "",
      id: item?.id ?? "",
      label: (item?.name ?? "").toString().trim(),
      img: item?.img ?? ""
    }))
    .filter((choice) => choice.label.length > 0 && (choice.uuid || choice.id))
    .sort((a, b) => a.label.localeCompare(b.label));
}

export function getWayangSpellhunterSelectedSpellName(item) {
  return getDictionaryString(item, WAYANG_SELECTED_SPELL_NAME_FLAG);
}

export async function setWayangSpellhunterSelectedSpellName(item, selectedSpellName) {
  if (!item || typeof item?.update !== "function") return "";
  const next = (selectedSpellName ?? "").toString().trim();
  await setDictionaryEntries(item, {
    [WAYANG_SELECTED_SPELL_NAME_FLAG]: next
  });
  return next;
}

async function isWayangSpellhunterItem(item) {
  if (!isFeatLikeItem(item)) return false;
  const reason = getFastMatchReason(item, {
    compendiumSource: WAYANG_SPELLHUNTER_COMPENDIUM_SOURCE,
    englishName: WAYANG_SPELLHUNTER_ENGLISH_NAME
  });
  if (reason) return true;
  if (isEnglishLanguage()) return false;
  const resolved = await resolveEnglishName(item?.name, { documentName: "Item", deepScanMode: "off" });
  return (resolved ?? "") === WAYANG_SPELLHUNTER_ENGLISH_NAME;
}

export function promptWayangSpellhunterSpellSelection(
  actor,
  {
    currentSpellName = "",
    title = WAYANG_SPELLHUNTER_ENGLISH_NAME,
    description = game.i18n.format("NAS.metamagic.dialogs.chooseSpellForName", {
      name: game.i18n.localize("NAS.metamagic.featureNames.wayangSpellhunter")
    })
  } = {}
) {
  const choices = getWayangSpellhunterSpellChoices(actor);
  if (!choices.length) {
    ui.notifications.warn(localizeMetamagic("warnings.noNonSpelllikeSpellsUpToThirdFound"));
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
        <label class="nas-wayang-spell-option" style="display:flex;align-items:center;gap:8px;padding:4px 0;">
          <input type="radio" name="wayangSpellhunterSpell" value="${value}" ${checked ? "checked" : ""} />
          ${img}
          <span>${safeLabel}</span>
        </label>
      `;
    })
    .join("");

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
            const selected = html.find('input[name="wayangSpellhunterSpell"]:checked').val();
            if (!selected) {
              resolve(null);
              return;
            }
            const selectedString = `${selected}`.trim();
            const selectedId = extractDocumentIdFromUuid(selectedString);
            const match = choices.find((choice) => {
              if (choice.uuid && choice.uuid === selectedString) return true;
              if (choice.id && choice.id === selectedString) return true;
              if (selectedId && choice.id && choice.id === selectedId) return true;
              return false;
            }) ?? null;
            if (!match) {
              resolve(null);
              return;
            }
            resolve({
              selectedSpellUuid: match.uuid || match.id || "",
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

export async function getWayangSpellhunterSource(actor, spellItem) {
  if (!actor || !spellItem || spellItem.type !== "spell") return null;
  const traitItem = await findActorTrait(actor, {
    compendiumSource: WAYANG_SPELLHUNTER_COMPENDIUM_SOURCE,
    englishName: WAYANG_SPELLHUNTER_ENGLISH_NAME
  });
  if (!traitItem) return null;

  const selectedSpellName = getWayangSpellhunterSelectedSpellName(traitItem);
  const spellChoices = getWayangSpellhunterSpellChoices(actor);
  const selectedChoice = findChoiceByStoredName(spellChoices, selectedSpellName);
  const selectedSpellUuid = selectedChoice?.uuid || selectedChoice?.id || "";
  const selectedSpellMatches = selectedSpellMatchesAction(spellItem, selectedSpellUuid);
  return {
    id: WAYANG_SPELLHUNTER_ID,
    label: traitItem?.name ?? WAYANG_SPELLHUNTER_ENGLISH_NAME,
    itemUuid: traitItem?.uuid ?? null,
    itemId: traitItem?.id ?? null,
    effectType: "wayangSpellhunterMinata",
    requiresMetamagicIntent: false,
    preserveSelectionWhenHidden: true,
    persistent: true,
    selectedSpellName,
    selectedSpellUuid,
    selectedSpellMatches,
    requiresSpellSelection: !selectedSpellUuid,
    metamagicSlotAdjustment: -1
  };
}

export async function prepareWayangSpellhunterContext(action, context) {
  const actor = context?.actor ?? action?.actor ?? action?.token?.actor ?? null;
  if (!actor || !action?.item) return;
  const source = await getWayangSpellhunterSource(actor, action.item);
  if (!source) return;

  const selectedTraitIds = Array.isArray(context?.traitOptions)
    ? context.traitOptions.map((value) => `${value}`)
    : [];
  const enabledForCast = selectedTraitIds.includes(WAYANG_SPELLHUNTER_ID);
  if (!enabledForCast || source.selectedSpellMatches !== true) return;

  context.featEffects ??= {};
  context.featEffects[WAYANG_SPELLHUNTER_ID] = {
    active: true,
    label: source.label,
    spellUuid: source.selectedSpellUuid,
    spellLabel: source.selectedSpellName || action?.item?.name || "",
    metamagicSlotAdjustment: -1
  };
}

export function getWayangSpellhunterSlotAdjustment(
  context,
  { hasAppliedMetamagic = false, timelessSoulActive = false } = {}
) {
  if (!hasAppliedMetamagic) return 0;
  if (timelessSoulActive) return 0;
  const adjustment = Number(
    context?.featEffects?.[WAYANG_SPELLHUNTER_ID]?.metamagicSlotAdjustment ?? 0
  );
  if (!Number.isFinite(adjustment)) return 0;
  return adjustment;
}

export function registerWayangSpellhunterItemHook() {
  if (wayangSpellhunterHookRegistered) return;
  wayangSpellhunterHookRegistered = true;
  Hooks.on("createItem", async (item, _options, userId) => {
    if (!item || userId !== game?.user?.id) return;
    if (!item?.parent || item.parent.documentName !== "Actor") return;
    if (!item?.isOwner) return;
    if (!(await isWayangSpellhunterItem(item))) return;
    if (getWayangSpellhunterSelectedSpellName(item)) return;
    const selection = await promptWayangSpellhunterSpellSelection(item.parent, {
      title: item?.name ?? WAYANG_SPELLHUNTER_ENGLISH_NAME
    });
    if (!selection?.selectedSpellName) return;
    await setWayangSpellhunterSelectedSpellName(item, selection.selectedSpellName);
  });
}
