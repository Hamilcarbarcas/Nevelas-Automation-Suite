import { MODULE } from "../../../../common/module.js";
import { contextHasSavingThrow, resolveFeatSaveDcBase } from "../../utils/saveDcUtils.js";
import { resolveEnglishName } from "../../utils/compendiumNameResolver.js";
import {
  findChoiceByStoredName,
  getDictionaryNumber,
  getDictionaryPrefixedStrings,
  getDictionaryString,
  setDictionaryEntries,
  setDictionaryPrefixedStrings
} from "../../utils/itemDictionarySelection.js";
import { isDurationEligibleForExtendSpell } from "../extendSpell.js";

const ELDRITCH_RESEARCHER_COMPENDIUM_SOURCE = "Compendium.pf-content.pf-feats.Item.h8G3iCn5v6N4SU0Q";
const ELDRITCH_RESEARCHER_ENGLISH_NAME = "Eldritch Researcher";
const ELDRITCH_RESEARCHER_ID = "eldritchResearcher";
const ELDRITCH_RESEARCHER_STATE_INCOMPLETE = "incomplete";
const ELDRITCH_RESEARCHER_STATE_COMPLETE = "complete";
const ELDRITCH_RESEARCHER_MODE_FLAG = "ERSM";
const ELDRITCH_RESEARCHER_SELECTED_SPELL_NAME_FLAG = "ERSSN";
const EXTENDED_SCRYING_COMPENDIUM_SOURCE = "Compendium.pf-content.pf-feats.Item.qIkjpBCtcgBLlKs8";
const EXTENDED_SCRYING_ENGLISH_NAME = "Extended Scrying";
const EXTENDED_SCRYING_ID = "extendedScrying";
const EXTENDED_SCRYING_FLAG_KEY = "feats.extendedScrying";
const SPELL_PERFECTION_COMPENDIUM_SOURCE = "Compendium.pf-content.pf-feats.Item.b19QTzVuFd0hlmrG";
const SPELL_PERFECTION_ENGLISH_NAME = "Spell Perfection";
const SPELL_PERFECTION_ID = "spellPerfection";
const SPELL_PERFECTION_SELECTED_SPELL_NAME_FLAG = "SPSSN";
const SPELL_PERFECTION_MIN_METAMAGIC_FEATS = 3;
const SPONTANEOUS_METAFOCUS_COMPENDIUM_SOURCE = "";
const SPONTANEOUS_METAFOCUS_ENGLISH_NAME = "Spontaneous Metafocus";
const SPONTANEOUS_METAFOCUS_ID = "spontaneousMetafocus";
const SPONTANEOUS_METAFOCUS_SELECTED_FLAG = "SMFSC";
const SPONTANEOUS_METAFOCUS_SPELL_FLAG_PREFIX = "SMFSS";

let eldritchResearcherHookRegistered = false;
let spellPerfectionHookRegistered = false;
let spontaneousMetafocusHookRegistered = false;

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

function isFeatItem(item) {
  const subType = item?.subType ?? item?.system?.subType;
  return item?.type === "feat" && (subType === "feat" || subType === "trait" || subType === "classFeat");
}

function isMetamagicFeatItem(item) {
  if (!isFeatItem(item)) return false;
  const tags = item?.system?.tags;
  if (!Array.isArray(tags)) return false;
  return tags.some((tag) => (tag ?? "").toString().trim().toLowerCase().includes("metamagic"));
}

function countMetamagicFeats(actor) {
  if (!actor?.items) return 0;
  return Array.from(actor.items).filter(isMetamagicFeatItem).length;
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

function getDefaultState() {
  return {
    enabled: false,
    mode: ELDRITCH_RESEARCHER_STATE_INCOMPLETE,
    selectedSpellUuid: "",
    selectedSpellLabel: "",
    selectedSpellName: ""
  };
}

function sanitizeState(rawState) {
  const raw = rawState && typeof rawState === "object" ? rawState : {};
  const mode = raw.mode === ELDRITCH_RESEARCHER_STATE_COMPLETE
    ? ELDRITCH_RESEARCHER_STATE_COMPLETE
    : ELDRITCH_RESEARCHER_STATE_INCOMPLETE;
  return {
    enabled: raw.enabled === true,
    mode,
    selectedSpellUuid: (raw.selectedSpellUuid ?? "").toString().trim(),
    selectedSpellLabel: (raw.selectedSpellLabel ?? "").toString().trim(),
    selectedSpellName: (raw.selectedSpellName ?? raw.selectedSpellLabel ?? "").toString().trim()
  };
}

function isSpellLikeSpellbook(spellItem) {
  return normalizeKey(spellItem?.system?.spellbook) === "spelllike";
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

export function getEldritchResearcherOptionId() {
  return ELDRITCH_RESEARCHER_ID;
}

export function getExtendedScryingOptionId() {
  return EXTENDED_SCRYING_ID;
}

export function getSpellPerfectionOptionId() {
  return SPELL_PERFECTION_ID;
}

export function getSpontaneousMetafocusOptionId() {
  return SPONTANEOUS_METAFOCUS_ID;
}

export function getEldritchResearcherStateModes() {
  return {
    incomplete: ELDRITCH_RESEARCHER_STATE_INCOMPLETE,
    complete: ELDRITCH_RESEARCHER_STATE_COMPLETE
  };
}

export async function getEldritchResearcherSource(actor) {
  if (!actor) return null;
  const featItem = await findActorFeat(actor, {
    compendiumSource: ELDRITCH_RESEARCHER_COMPENDIUM_SOURCE,
    englishName: ELDRITCH_RESEARCHER_ENGLISH_NAME
  });
  if (!featItem) return null;
  return {
    id: ELDRITCH_RESEARCHER_ID,
    label: featItem?.name ?? ELDRITCH_RESEARCHER_ENGLISH_NAME,
    itemUuid: featItem?.uuid ?? null,
    persistent: true
  };
}

export async function getExtendedScryingSource(actor) {
  if (!actor) return null;
  const featItem = await findActorFeat(actor, {
    compendiumSource: EXTENDED_SCRYING_COMPENDIUM_SOURCE,
    englishName: EXTENDED_SCRYING_ENGLISH_NAME
  });
  if (!featItem) return null;
  return {
    id: EXTENDED_SCRYING_ID,
    label: featItem?.name ?? EXTENDED_SCRYING_ENGLISH_NAME,
    itemUuid: featItem?.uuid ?? null,
    persistent: true
  };
}

export async function getSpellPerfectionSource(actor) {
  if (!actor) return null;
  if (countMetamagicFeats(actor) < SPELL_PERFECTION_MIN_METAMAGIC_FEATS) return null;
  const featItem = await findActorFeat(actor, {
    compendiumSource: SPELL_PERFECTION_COMPENDIUM_SOURCE,
    englishName: SPELL_PERFECTION_ENGLISH_NAME
  });
  if (!featItem) return null;
  return {
    id: SPELL_PERFECTION_ID,
    label: featItem?.name ?? SPELL_PERFECTION_ENGLISH_NAME,
    itemUuid: featItem?.uuid ?? null,
    persistent: true
  };
}

export async function getSpontaneousMetafocusSource(actor) {
  if (!actor) return null;
  const featItem = await findActorFeat(actor, {
    compendiumSource: SPONTANEOUS_METAFOCUS_COMPENDIUM_SOURCE,
    englishName: SPONTANEOUS_METAFOCUS_ENGLISH_NAME
  });
  if (!featItem) return null;
  return {
    id: SPONTANEOUS_METAFOCUS_ID,
    label: featItem?.name ?? SPONTANEOUS_METAFOCUS_ENGLISH_NAME,
    itemUuid: featItem?.uuid ?? null,
    persistent: false
  };
}

export async function getEldritchResearcherState(actor) {
  const featItem = await findActorFeat(actor, {
    compendiumSource: ELDRITCH_RESEARCHER_COMPENDIUM_SOURCE,
    englishName: ELDRITCH_RESEARCHER_ENGLISH_NAME
  });
  const selectedSpellName = getDictionaryString(featItem, ELDRITCH_RESEARCHER_SELECTED_SPELL_NAME_FLAG);
  return sanitizeState({
    enabled: false,
    mode: getDictionaryString(featItem, ELDRITCH_RESEARCHER_MODE_FLAG) || ELDRITCH_RESEARCHER_STATE_INCOMPLETE,
    selectedSpellLabel: selectedSpellName,
    selectedSpellName
  });
}

export async function setEldritchResearcherState(actor, nextState = {}) {
  const featItem = await findActorFeat(actor, {
    compendiumSource: ELDRITCH_RESEARCHER_COMPENDIUM_SOURCE,
    englishName: ELDRITCH_RESEARCHER_ENGLISH_NAME
  });
  if (!featItem) return getDefaultState();
  const current = await getEldritchResearcherState(actor);
  const merged = sanitizeState({ ...current, ...nextState });
  await setDictionaryEntries(featItem, {
    [ELDRITCH_RESEARCHER_MODE_FLAG]: merged.mode,
    [ELDRITCH_RESEARCHER_SELECTED_SPELL_NAME_FLAG]: merged.selectedSpellName || merged.selectedSpellLabel || ""
  });
  return merged;
}

export function getExtendedScryingState(actor) {
  const raw = actor?.getFlag?.(MODULE.ID, EXTENDED_SCRYING_FLAG_KEY);
  return raw === true;
}

export async function setExtendedScryingState(actor, enabled) {
  if (!actor || typeof actor?.setFlag !== "function") return false;
  const nextValue = enabled === true;
  await actor.setFlag(MODULE.ID, EXTENDED_SCRYING_FLAG_KEY, nextValue);
  return nextValue;
}

function sanitizeSpellPerfectionState(rawState) {
  const raw = rawState && typeof rawState === "object" ? rawState : {};
  return {
    enabled: raw.enabled === true,
    selectedSpellUuid: (raw.selectedSpellUuid ?? "").toString().trim(),
    selectedSpellLabel: (raw.selectedSpellLabel ?? "").toString().trim(),
    selectedSpellName: (raw.selectedSpellName ?? raw.selectedSpellLabel ?? "").toString().trim()
  };
}

function getDefaultSpellPerfectionState() {
  return {
    enabled: false,
    selectedSpellUuid: "",
    selectedSpellLabel: "",
    selectedSpellName: ""
  };
}

export async function getSpellPerfectionState(actor) {
  const featItem = await findActorFeat(actor, {
    compendiumSource: SPELL_PERFECTION_COMPENDIUM_SOURCE,
    englishName: SPELL_PERFECTION_ENGLISH_NAME
  });
  const selectedSpellName = getDictionaryString(featItem, SPELL_PERFECTION_SELECTED_SPELL_NAME_FLAG);
  return sanitizeSpellPerfectionState({
    enabled: false,
    selectedSpellLabel: selectedSpellName,
    selectedSpellName
  });
}

export async function setSpellPerfectionState(actor, nextState = {}) {
  const featItem = await findActorFeat(actor, {
    compendiumSource: SPELL_PERFECTION_COMPENDIUM_SOURCE,
    englishName: SPELL_PERFECTION_ENGLISH_NAME
  });
  if (!featItem) return getDefaultSpellPerfectionState();
  const current = await getSpellPerfectionState(actor);
  const merged = sanitizeSpellPerfectionState({ ...current, ...nextState });
  await setDictionaryEntries(featItem, {
    [SPELL_PERFECTION_SELECTED_SPELL_NAME_FLAG]: merged.selectedSpellName || merged.selectedSpellLabel || ""
  });
  return merged;
}

async function isEldritchResearcherItem(item) {
  if (!isFeatItem(item)) return false;
  const reason = getFeatFastMatchReason(item, {
    compendiumSource: ELDRITCH_RESEARCHER_COMPENDIUM_SOURCE,
    englishName: ELDRITCH_RESEARCHER_ENGLISH_NAME
  });
  if (reason) return true;
  if (isEnglishLanguage()) return false;
  const resolved = await resolveEnglishName(item?.name, { documentName: "Item", deepScanMode: "off" });
  return (resolved ?? "") === ELDRITCH_RESEARCHER_ENGLISH_NAME;
}

async function isSpellPerfectionItem(item) {
  if (!isFeatItem(item)) return false;
  const reason = getFeatFastMatchReason(item, {
    compendiumSource: SPELL_PERFECTION_COMPENDIUM_SOURCE,
    englishName: SPELL_PERFECTION_ENGLISH_NAME
  });
  if (reason) return true;
  if (isEnglishLanguage()) return false;
  const resolved = await resolveEnglishName(item?.name, { documentName: "Item", deepScanMode: "off" });
  return (resolved ?? "") === SPELL_PERFECTION_ENGLISH_NAME;
}

async function isSpontaneousMetafocusItem(item) {
  if (!isFeatItem(item)) return false;
  const reason = getFeatFastMatchReason(item, {
    compendiumSource: SPONTANEOUS_METAFOCUS_COMPENDIUM_SOURCE,
    englishName: SPONTANEOUS_METAFOCUS_ENGLISH_NAME
  });
  if (reason) return true;
  if (isEnglishLanguage()) return (item?.name ?? "") === SPONTANEOUS_METAFOCUS_ENGLISH_NAME;
  const resolved = await resolveEnglishName(item?.name, { documentName: "Item", deepScanMode: "off" });
  return (resolved ?? "") === SPONTANEOUS_METAFOCUS_ENGLISH_NAME;
}

function promptSpontaneousMetafocusInitialSelection(actor, { selectedCount = 1, title = SPONTANEOUS_METAFOCUS_ENGLISH_NAME } = {}) {
  const choices = getEldritchResearcherSpellChoices(actor);
  if (!choices.length) {
    ui.notifications.warn(localizeMetamagic("warnings.noNonSpelllikeSpellsFound"));
    return Promise.resolve(null);
  }
  const maxSelections = Math.max(1, Number(selectedCount) || 1);
  const useCheckboxes = maxSelections > 1;
  const inputType = useCheckboxes ? "checkbox" : "radio";
  const inputName = "spontaneousMetafocusInitialChoice";
  const cards = choices
    .map((choice, index) => {
      const value = (choice.label ?? "").toString().trim();
      if (!value) return "";
      const img = choice.img ? `<img src="${choice.img}" style="width:32px;height:32px;border-radius:4px;" />` : "";
      const safeLabel = foundry.utils.escapeHTML(choice.label ?? `Spell ${index + 1}`);
      return `
        <label class="nas-spontaneous-metafocus-spell-option" style="display:flex;align-items:center;gap:8px;padding:4px 0;">
          <input type="${inputType}" name="${inputName}" value="${value}" />
          ${img}
          <span>${safeLabel}</span>
        </label>
      `;
    })
    .join("");
  const description = useCheckboxes
    ? formatMetamagic("dialogs.chooseSpellsForName", { count: maxSelections, name: title })
    : formatMetamagic("dialogs.chooseSpellForName", { name: title });
  return new Promise((resolve) => {
    new Dialog({
      title,
      content: `
        <form>
          <p style="margin:0 0 8px 0;">${foundry.utils.escapeHTML(description)}</p>
          <div style="max-height:320px;overflow-y:auto;border:1px solid #888;border-radius:4px;padding:6px 8px;">
            ${cards}
          </div>
        </form>
      `,
      buttons: {
        ok: {
          label: game.i18n.localize("NAS.common.buttons.select"),
          callback: (html) => {
            const values = html
              .find(`input[name="${inputName}"]:checked`)
              .map((_, el) => (el?.value ?? "").toString().trim())
              .get()
              .filter(Boolean);
            if (values.length !== maxSelections) {
              resolve(null);
              return;
            }
            resolve(values);
          }
        },
        cancel: {
          label: game.i18n.localize("NAS.common.buttons.cancel"),
          callback: () => resolve(null)
        }
      },
      default: "ok",
      close: () => resolve(null),
      render: (html) => {
        if (!useCheckboxes) return;
        html.find(`input[name="${inputName}"]`).on("change", (event) => {
          const checked = html.find(`input[name="${inputName}"]:checked`).length;
          if (checked > maxSelections) {
            event.currentTarget.checked = false;
          }
        });
      }
    }).render(true);
  });
}

export function registerEldritchResearcherItemHook() {
  if (eldritchResearcherHookRegistered) return;
  eldritchResearcherHookRegistered = true;
  Hooks.on("createItem", async (item, _options, userId) => {
    if (!item || userId !== game?.user?.id) return;
    if (!item?.parent || item.parent.documentName !== "Actor") return;
    if (!item?.isOwner) return;
    if (!(await isEldritchResearcherItem(item))) return;
    const selectedSpellName = getDictionaryString(item, ELDRITCH_RESEARCHER_SELECTED_SPELL_NAME_FLAG);
    if (selectedSpellName) return;
    const selection = await promptEldritchResearcherSpellSelection(item.parent, {
      title: item?.name ?? ELDRITCH_RESEARCHER_ENGLISH_NAME,
      includeModeToggle: true
    });
    if (!selection?.selectedSpellName) return;
    await setEldritchResearcherState(item.parent, {
      enabled: false,
      mode: selection.selectedMode ?? ELDRITCH_RESEARCHER_STATE_INCOMPLETE,
      selectedSpellName: selection.selectedSpellName,
      selectedSpellLabel: selection.selectedSpellLabel
    });
  });
}

export function registerSpellPerfectionItemHook() {
  if (spellPerfectionHookRegistered) return;
  spellPerfectionHookRegistered = true;
  Hooks.on("createItem", async (item, _options, userId) => {
    if (!item || userId !== game?.user?.id) return;
    if (!item?.parent || item.parent.documentName !== "Actor") return;
    if (!item?.isOwner) return;
    if (!(await isSpellPerfectionItem(item))) return;
    const selectedSpellName = getDictionaryString(item, SPELL_PERFECTION_SELECTED_SPELL_NAME_FLAG);
    if (selectedSpellName) return;
    const selection = await promptEldritchResearcherSpellSelection(item.parent, {
      title: item?.name ?? SPELL_PERFECTION_ENGLISH_NAME,
      description: formatMetamagic("dialogs.chooseSpellForName", {
        name: item?.name ?? SPELL_PERFECTION_ENGLISH_NAME
      })
    });
    if (!selection?.selectedSpellName) return;
    await setSpellPerfectionState(item.parent, {
      enabled: false,
      selectedSpellName: selection.selectedSpellName,
      selectedSpellLabel: selection.selectedSpellLabel
    });
  });
}

export function registerSpontaneousMetafocusItemHook() {
  if (spontaneousMetafocusHookRegistered) return;
  spontaneousMetafocusHookRegistered = true;
  Hooks.on("createItem", async (item, _options, userId) => {
    if (!item || userId !== game?.user?.id) return;
    if (!item?.parent || item.parent.documentName !== "Actor") return;
    if (!item?.isOwner) return;
    if (!(await isSpontaneousMetafocusItem(item))) return;
    const selectedNames = getDictionaryPrefixedStrings(item, SPONTANEOUS_METAFOCUS_SPELL_FLAG_PREFIX);
    if (selectedNames.length > 0) return;
    const selectedCount = Math.max(1, Number(getDictionaryNumber(item, SPONTANEOUS_METAFOCUS_SELECTED_FLAG, 1) || 1));
    const chosen = await promptSpontaneousMetafocusInitialSelection(item.parent, {
      selectedCount,
      title: item?.name ?? SPONTANEOUS_METAFOCUS_ENGLISH_NAME
    });
    if (!chosen?.length) return;
    await setDictionaryEntries(item, {
      [SPONTANEOUS_METAFOCUS_SELECTED_FLAG]: String(selectedCount)
    });
    await setDictionaryPrefixedStrings(item, SPONTANEOUS_METAFOCUS_SPELL_FLAG_PREFIX, chosen);
  });
}

export function getEldritchResearcherSpellChoices(actor) {
  if (!actor?.items) return [];
  return Array.from(actor.items)
    .filter((item) => item?.type === "spell")
    .filter((item) => !isSpellLikeSpellbook(item))
    .map((item) => ({
      uuid: item?.uuid ?? "",
      id: item?.id ?? "",
      label: item?.name ?? "Spell",
      img: item?.img ?? "",
      spellbook: item?.system?.spellbook ?? "",
      level: Number(item?.system?.level ?? 0)
    }))
    .filter((entry) => entry.uuid || entry.id)
    .sort((a, b) => a.label.localeCompare(b.label));
}

export function resolveSpellLabelFromChoices(choices, selectedSpellUuid, fallback = "") {
  const selectedId = extractDocumentIdFromUuid(selectedSpellUuid);
  const match = choices.find((entry) => {
    if (entry.uuid && entry.uuid === selectedSpellUuid) return true;
    if (selectedId && entry.id && entry.id === selectedId) return true;
    return false;
  });
  return match?.label ?? fallback;
}

export function promptEldritchResearcherSpellSelection(
  actor,
  {
    currentSpellUuid = "",
    currentMode = ELDRITCH_RESEARCHER_STATE_INCOMPLETE,
    includeModeToggle = false,
    title = ELDRITCH_RESEARCHER_ENGLISH_NAME,
    description = game.i18n.format("NAS.metamagic.dialogs.chooseSpellForName", {
      name: game.i18n.localize("NAS.metamagic.featureNames.eldritchResearcher")
    })
  } = {}
) {
  const choices = getEldritchResearcherSpellChoices(actor);
  if (!choices.length) {
    ui.notifications.warn(localizeMetamagic("warnings.noNonSpelllikeSpellsFound"));
    return Promise.resolve(null);
  }

  const selectedId = extractDocumentIdFromUuid(currentSpellUuid);
  const cards = choices
    .map((choice, index) => {
      const checked = choice.uuid === currentSpellUuid || (selectedId && choice.id === selectedId);
      const value = choice.uuid || choice.id;
      const img = choice.img ? `<img src="${choice.img}" style="width:32px;height:32px;border-radius:4px;" />` : "";
      const safeLabel = foundry.utils.escapeHTML(choice.label ?? `Spell ${index + 1}`);
      return `
        <label class="nas-eldritch-spell-option" style="display:flex;align-items:center;gap:8px;padding:4px 0;">
          <input type="radio" name="eldritchSpell" value="${value}" ${checked ? "checked" : ""} />
          ${img}
          <span>${safeLabel}</span>
        </label>
      `;
    })
    .join("");

  const content = `
    <form>
      <p style="margin:0 0 8px 0;">${foundry.utils.escapeHTML(description)}</p>
      ${includeModeToggle ? `
      <div class="form-group" style="margin:0 0 8px 0;">
        <label style="display:flex;align-items:center;gap:6px;">
          <input type="checkbox" name="eldritchResearcherComplete" ${currentMode === ELDRITCH_RESEARCHER_STATE_COMPLETE ? "checked" : ""} />
          ${localizeMetamagic("eldritchResearcher.storyFeatComplete")}
        </label>
      </div>
      ` : ""}
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
            const selected = html.find('input[name="eldritchSpell"]:checked').val();
            if (!selected) {
              resolve(null);
              return;
            }
            const selectedString = `${selected}`;
            const label = resolveSpellLabelFromChoices(choices, selectedString, "");
            const selectedMode = includeModeToggle && html.find('input[name="eldritchResearcherComplete"]').is(":checked")
              ? ELDRITCH_RESEARCHER_STATE_COMPLETE
              : ELDRITCH_RESEARCHER_STATE_INCOMPLETE;
            resolve({
              selectedSpellUuid: selectedString,
              selectedSpellLabel: label,
              selectedSpellName: label,
              selectedMode
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

export async function prepareEldritchResearcherContext(action, context) {
  const actor = context?.actor ?? action?.actor ?? action?.token?.actor ?? null;
  if (!actor || !action?.item) return;
  const source = await getEldritchResearcherSource(actor);
  if (!source) return;

  const optionState = sanitizeState(
    context?.featOptions?.[ELDRITCH_RESEARCHER_ID] ?? await getEldritchResearcherState(actor)
  );
  if (!optionState.enabled) return;
  let selectedSpellMatches = false;
  if (optionState.selectedSpellUuid) {
    selectedSpellMatches = selectedSpellMatchesAction(action.item, optionState.selectedSpellUuid);
  } else if (optionState.selectedSpellName) {
    const spellChoices = getEldritchResearcherSpellChoices(actor);
    const selectedChoice = findChoiceByStoredName(spellChoices, optionState.selectedSpellName);
    const selectedSpellUuid = selectedChoice?.uuid || selectedChoice?.id || "";
    if (selectedSpellUuid) {
      optionState.selectedSpellUuid = selectedSpellUuid;
      optionState.selectedSpellLabel = selectedChoice?.label ?? optionState.selectedSpellName;
      selectedSpellMatches = selectedSpellMatchesAction(action.item, selectedSpellUuid);
    }
  }
  if (!selectedSpellMatches) return;

  const mode = optionState.mode === ELDRITCH_RESEARCHER_STATE_COMPLETE
    ? ELDRITCH_RESEARCHER_STATE_COMPLETE
    : ELDRITCH_RESEARCHER_STATE_INCOMPLETE;

  context.featEffects ??= {};
  context.featEffects[ELDRITCH_RESEARCHER_ID] = {
    active: true,
    mode,
    label: source.label,
    spellUuid: optionState.selectedSpellUuid,
    spellLabel: optionState.selectedSpellLabel || action?.item?.name || "",
    casterLevelBonus: 1,
    saveDcBonus: mode === ELDRITCH_RESEARCHER_STATE_COMPLETE ? 1 : 0,
    metamagicSlotAdjustment: mode === ELDRITCH_RESEARCHER_STATE_COMPLETE ? -1 : 0
  };

  const currentCl = Number(action?.shared?.rollData?.cl ?? 0);
  if (Number.isFinite(currentCl)) {
    action.shared.rollData.cl = currentCl + 1;
  }
}

export function getEldritchResearcherSlotAdjustment(context, { hasAppliedMetamagic = false } = {}) {
  if (!hasAppliedMetamagic) return 0;
  const adjustment = Number(
    context?.featEffects?.[ELDRITCH_RESEARCHER_ID]?.metamagicSlotAdjustment ?? 0
  );
  if (!Number.isFinite(adjustment)) return 0;
  return adjustment;
}

export function applyEldritchResearcherPostMetamagic(action, context) {
  const effect = context?.featEffects?.[ELDRITCH_RESEARCHER_ID];
  if (!effect?.active) return;
  if (effect?.mode !== ELDRITCH_RESEARCHER_STATE_COMPLETE) return;
  if (!contextHasSavingThrow(context, action)) return;

  const currentDc = resolveFeatSaveDcBase(action, context);
  if (!Number.isFinite(currentDc)) return;
  const nextDc = currentDc + 1;

  context.save ??= {};
  const existingBase = Number(context.save.baseDc);
  if (!Number.isFinite(existingBase)) {
    context.save.baseDc = currentDc;
  }
  context.save.dc = nextDc;
  const evaluatedTotal = Number(context?.save?.evaluated?.total);
  if (Number.isFinite(evaluatedTotal)) {
    context.save.evaluated.total = evaluatedTotal + 1;
  } else {
    context.save.evaluated ??= {};
    context.save.evaluated.total = nextDc;
  }

  action.shared ??= {};
  action.shared.saveDC = nextDc;
}

function hasScryingSubschool(item) {
  const subschool = item?.system?.subschool;
  if (subschool && typeof subschool === "object" && !Array.isArray(subschool)) {
    const candidates = [];
    if (Array.isArray(subschool.base)) candidates.push(...subschool.base);
    if (Array.isArray(subschool.names)) candidates.push(...subschool.names);
    if (Array.isArray(subschool.total)) candidates.push(...subschool.total);
    return candidates.some((entry) => normalizeKey(entry) === "scrying");
  }
  if (Array.isArray(subschool)) {
    return subschool.some((entry) => normalizeKey(entry) === "scrying");
  }
  if (typeof subschool === "string") {
    return normalizeKey(subschool) === "scrying";
  }
  return false;
}

function applyExtendedScryingDuration(context) {
  if (!context?.duration) return false;
  const hasExtendAlready = Array.isArray(context?.metamagicNames)
    && context.metamagicNames.includes("Extend Spell");
  if (hasExtendAlready) return false;
  if (!isDurationEligibleForExtendSpell(context.duration)) return false;

  const evaluatedBase = Number(context.duration.evaluated?.total);
  if (Number.isFinite(evaluatedBase)) {
    const total = evaluatedBase * 2;
    context.duration.value = String(total);
    context.duration.evaluated = {
      ...(context.duration.evaluated ?? {}),
      total
    };
    return true;
  }

  const rawValue = Number(context.duration.value ?? 0);
  if (!Number.isFinite(rawValue) || rawValue <= 0) return false;
  const doubled = rawValue * 2;
  context.duration.value = String(doubled);
  context.duration.evaluated = {
    ...(context.duration.evaluated ?? {}),
    total: doubled
  };
  return true;
}

function applyExtendedScryingCastTime(context) {
  if (!context?.activation) return false;
  const type = normalizeKey(context.activation.type);
  const rawCost = Number(context.activation.cost);
  const cost = Number.isFinite(rawCost) && rawCost > 0 ? rawCost : 1;
  let changed = false;

  if (type === "minute" && cost >= 10) {
    const nextCost = cost / 2;
    context.activation.cost = nextCost;
    context.activation.type = "minute";
    context.activation.unchained ??= {};
    context.activation.unchained.cost = nextCost;
    context.activation.unchained.type = "minute";
    changed = true;
  } else if (type === "hour") {
    const nextCost = cost * 30;
    context.activation.cost = nextCost;
    context.activation.type = "minute";
    context.activation.unchained ??= {};
    context.activation.unchained.cost = nextCost;
    context.activation.unchained.type = "minute";
    changed = true;
  }

  return changed;
}

export async function prepareExtendedScryingContext(action, context) {
  const actor = context?.actor ?? action?.actor ?? action?.token?.actor ?? null;
  if (!actor || !action?.item) return;
  const source = await getExtendedScryingSource(actor);
  if (!source) return;

  const formEnabled = context?.featOptions?.[EXTENDED_SCRYING_ID]?.enabled;
  const enabled = formEnabled === true || (formEnabled === undefined && getExtendedScryingState(actor) === true);
  if (!enabled) return;
  if (action.item.type !== "spell") return;
  if (!hasScryingSubschool(action.item)) return;

  const durationApplied = applyExtendedScryingDuration(context);
  const castTimeApplied = applyExtendedScryingCastTime(context);
  if (!durationApplied && !castTimeApplied) return;

  context.featEffects ??= {};
  context.featEffects[EXTENDED_SCRYING_ID] = {
    active: true,
    label: source.label,
    durationApplied,
    castTimeApplied
  };
}

export async function prepareSpellPerfectionContext(action, context) {
  const actor = context?.actor ?? action?.actor ?? action?.token?.actor ?? null;
  if (!actor || !action?.item) return;
  const source = await getSpellPerfectionSource(actor);
  if (!source) return;

  const optionState = sanitizeSpellPerfectionState(
    context?.featOptions?.[SPELL_PERFECTION_ID] ?? await getSpellPerfectionState(actor)
  );
  if (!optionState.enabled) return;
  let selectedSpellMatches = false;
  if (optionState.selectedSpellUuid) {
    selectedSpellMatches = selectedSpellMatchesAction(action.item, optionState.selectedSpellUuid);
  } else if (optionState.selectedSpellName) {
    const spellChoices = getEldritchResearcherSpellChoices(actor);
    const selectedChoice = findChoiceByStoredName(spellChoices, optionState.selectedSpellName);
    const selectedSpellUuid = selectedChoice?.uuid || selectedChoice?.id || "";
    if (selectedSpellUuid) {
      optionState.selectedSpellUuid = selectedSpellUuid;
      optionState.selectedSpellLabel = selectedChoice?.label ?? optionState.selectedSpellName;
      selectedSpellMatches = selectedSpellMatchesAction(action.item, selectedSpellUuid);
    }
  }
  if (!selectedSpellMatches) return;

  context.featEffects ??= {};
  context.featEffects[SPELL_PERFECTION_ID] = {
    active: true,
    label: source.label,
    spellUuid: optionState.selectedSpellUuid,
    spellLabel: optionState.selectedSpellLabel || action?.item?.name || ""
  };
}

function sanitizeSpontaneousMetafocusState(rawState) {
  const raw = rawState && typeof rawState === "object" ? rawState : {};
  const selectedSpellUuids = Array.isArray(raw.selectedSpellUuids)
    ? raw.selectedSpellUuids
    : [];
  return {
    enabled: raw.enabled === true,
    selectedSpellUuids: Array.from(
      new Set(
        selectedSpellUuids
          .map((value) => (value ?? "").toString().trim())
          .filter(Boolean)
      )
    )
  };
}

function selectedSpellListMatchesAction(actionItem, selectedSpellUuids = []) {
  if (!Array.isArray(selectedSpellUuids) || !selectedSpellUuids.length) return false;
  return selectedSpellUuids.some((uuid) => selectedSpellMatchesAction(actionItem, uuid));
}

export async function prepareSpontaneousMetafocusContext(action, context) {
  const actor = context?.actor ?? action?.actor ?? action?.token?.actor ?? null;
  if (!actor || !action?.item) return;
  const source = await getSpontaneousMetafocusSource(actor);
  if (!source) return;

  const optionState = sanitizeSpontaneousMetafocusState(
    context?.featOptions?.[SPONTANEOUS_METAFOCUS_ID]
  );
  if (!optionState.enabled) return;
  if (!selectedSpellListMatchesAction(action.item, optionState.selectedSpellUuids)) return;

  context.featEffects ??= {};
  context.featEffects[SPONTANEOUS_METAFOCUS_ID] = {
    active: true,
    label: source.label
  };
}

export function getSpellPerfectionStatus(context) {
  const effect = context?.featEffects?.[SPELL_PERFECTION_ID];
  return {
    enabled: effect?.active === true,
    label: effect?.label ?? SPELL_PERFECTION_ENGLISH_NAME
  };
}

export function getSpontaneousMetafocusStatus(context) {
  const effect = context?.featEffects?.[SPONTANEOUS_METAFOCUS_ID];
  return {
    enabled: effect?.active === true,
    label: effect?.label ?? SPONTANEOUS_METAFOCUS_ENGLISH_NAME
  };
}
