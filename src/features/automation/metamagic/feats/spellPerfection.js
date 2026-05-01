import {
  getDictionaryString
} from "../../utils/itemDictionarySelection.js";

const SPELL_PERFECTION_ID = "spellPerfection";
const SPELL_PERFECTION_COMPENDIUM_SOURCE = "Compendium.pf-content.pf-feats.Item.b19QTzVuFd0hlmrG";
const SPELL_PERFECTION_SELECTED_SPELL_NAME_FLAG = "SPSSN";
const CKL_MODULE_ID = "ckl-roll-bonuses";
const CKL_ACTION_USE_ALTER_HOOK = `${CKL_MODULE_ID}_actionUseAlterRollData`;
const DAMAGE_ELEMENT_KEYS = new Set(["acid", "cold", "electric", "fire"]);

let hooksRegistered = false;

function isFeatSourceItem(item) {
  if (!item || item.type !== "feat") return false;
  const subType = item?.subType ?? item?.system?.subType;
  return subType === "feat";
}

function extractDocumentIdFromUuid(uuid) {
  const value = (uuid ?? "").toString().trim();
  if (!value) return "";
  const parts = value.split(".");
  return parts.length ? (parts[parts.length - 1] ?? "") : "";
}

function selectedSpellMatchesAction(actionItem, selectedSpellUuid) {
  if (!actionItem || !selectedSpellUuid) return false;
  const selectedId = extractDocumentIdFromUuid(selectedSpellUuid);
  const actionUuid = (actionItem?.uuid ?? "").toString().trim();
  const actionId = (actionItem?.id ?? "").toString().trim();
  if (actionUuid && actionUuid === selectedSpellUuid) return true;
  if (selectedId && actionId && actionId === selectedId) return true;
  return false;
}

function normalizeKey(value) {
  return (value ?? "").toString().trim().toLowerCase();
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

function actionMatchesSelectedSpellName(actionItem, selectedSpellName) {
  const selected = normalizeKey(selectedSpellName);
  if (!actionItem || !selected) return false;
  const candidates = [
    normalizeKey(actionItem?.name),
    normalizeKey(getBabeleOriginalName(actionItem))
  ].filter(Boolean);
  return candidates.includes(selected);
}

function getActiveSpellPerfectionStateFromActionUse(actionUse) {
  const context = actionUse?.shared?.nasSpellContext;
  const effect = context?.featEffects?.[SPELL_PERFECTION_ID];
  if (effect?.active !== true) return null;
  const selectedSpellUuid = (effect?.spellUuid ?? "").toString().trim();
  if (!selectedSpellUuid) return null;
  if (!selectedSpellMatchesAction(actionUse?.item, selectedSpellUuid)) return null;
  return effect;
}

function getActiveSpellPerfectionStateFromActorAndAction(actor, actionUseLike) {
  if (!actor) return null;
  const featItem = Array.from(actor.items ?? []).find(
    (item) => (item?._stats?.compendiumSource ?? "") === SPELL_PERFECTION_COMPENDIUM_SOURCE
  ) ?? null;
  if (!featItem) return null;
  const selectedSpellName = getDictionaryString(featItem, SPELL_PERFECTION_SELECTED_SPELL_NAME_FLAG, { normalize: true });
  if (!selectedSpellName) return null;
  if (!actionMatchesSelectedSpellName(actionUseLike?.item, selectedSpellName)) return null;
  return { selectedSpellName };
}

function getCklApi() {
  return game?.modules?.get?.(CKL_MODULE_ID)?.api ?? null;
}

function getRollBonusesGlobal() {
  return globalThis?.RollBonuses ?? null;
}

function toLowerStringArray(values) {
  return (Array.isArray(values) ? values : [])
    .map((value) => (value ?? "").toString().trim().toLowerCase())
    .filter(Boolean);
}

function arrayIntersects(left, right) {
  const leftValues = toLowerStringArray(left);
  const rightValues = new Set(toLowerStringArray(right));
  if (!leftValues.length || !rightValues.size) return false;
  return leftValues.some((value) => rightValues.has(value));
}

function getSpellSchoolKeys(actionUse) {
  const raw = actionUse?.item?.system?.school;
  if (Array.isArray(raw)) return toLowerStringArray(raw);
  const school = (raw ?? "").toString().trim().toLowerCase();
  return school ? [school] : [];
}

function getActionBaseTypes(actionUse) {
  return toLowerStringArray(actionUse?.item?.system?.baseTypes);
}

function getSpellElementKeys(actionUse) {
  const out = new Set();
  const descriptors = Array.isArray(actionUse?.item?.system?.descriptors?.total)
    ? actionUse.item.system.descriptors.total
    : [];
  for (const rawDescriptor of descriptors) {
    const tokens = (rawDescriptor ?? "")
      .toString()
      .toLowerCase()
      .split(/,|\bor\b/)
      .map((part) => part.trim())
      .filter(Boolean);
    for (const token of tokens) {
      if (DAMAGE_ELEMENT_KEYS.has(token)) out.add(token);
    }
  }

  const damageParts = actionUse?.action?.data?.damage?.parts;
  if (Array.isArray(damageParts)) {
    for (const part of damageParts) {
      const types = Array.isArray(part?.types) ? part.types : [];
      for (const type of types) {
        const normalized = (type ?? "").toString().trim().toLowerCase();
        if (DAMAGE_ELEMENT_KEYS.has(normalized)) out.add(normalized);
      }
    }
  }

  return [...out];
}

function getSpecificBonusType(key) {
  const map = getRollBonusesGlobal()?.specificBonusTypeMap;
  if (!map || typeof map !== "object") return null;
  return map[key] ?? null;
}

function evaluateItemFormulaAsNumber(item, flagKey, actor) {
  const raw = item?.getFlag?.(CKL_MODULE_ID, flagKey);
  const formula = (raw ?? "").toString().trim();
  if (!formula) return 0;
  const direct = Number(formula);
  if (Number.isFinite(direct)) return direct;
  try {
    const rollData = actor?.getRollData?.() ?? {};
    const total = Roll?.defaultImplementation?.safeRollSync?.(formula, rollData)?.total;
    return Number.isFinite(total) ? Number(total) : 0;
  } catch (_err) {
    return 0;
  }
}

function getTargetedFeatDcBonusFromUpdate(actionUse) {
  const handleBonusesFor = getCklApi()?.utils?.handleBonusesFor;
  if (typeof handleBonusesFor !== "function") return 0;
  let total = 0;
  handleBonusesFor(actionUse, (bonusType, sourceItem) => {
    if (!isFeatSourceItem(sourceItem) || !isItemActive(sourceItem)) return;
    if (typeof bonusType?.updateItemActionRollData !== "function") return;
    const scratch = { dcBonus: 0 };
    bonusType.updateItemActionRollData(sourceItem, actionUse?.action, scratch);
    const delta = Number(scratch.dcBonus ?? 0);
    if (!Number.isFinite(delta) || delta === 0) return;
    total += delta;
  });
  return total;
}

function getTargetedFeatClBonusFromUpdate(actionUse) {
  const handleBonusesFor = getCklApi()?.utils?.handleBonusesFor;
  if (typeof handleBonusesFor !== "function") return 0;
  let total = 0;
  handleBonusesFor(actionUse, (bonusType, sourceItem) => {
    if (!isFeatSourceItem(sourceItem) || !isItemActive(sourceItem)) return;
    if (typeof bonusType?.updateItemActionRollData !== "function") return;
    const scratch = { cl: 0 };
    bonusType.updateItemActionRollData(sourceItem, actionUse?.action, scratch);
    const delta = Number(scratch.cl ?? 0);
    if (!Number.isFinite(delta) || delta === 0) return;
    total += delta;
  });
  return total;
}

function getSpellFocusFeatDcBonus(actor, actionUse) {
  if (!actor || actionUse?.item?.type !== "spell") return 0;
  const spellFocus = getSpecificBonusType("spell-focus");
  const spellFocusGreater = getSpecificBonusType("spell-focus-greater");
  const spellFocusMythic = getSpecificBonusType("spell-focus-mythic");
  const schools = getSpellSchoolKeys(actionUse);
  if (!schools.length) return 0;
  const focused = arrayIntersects(schools, spellFocus?.getFocusedSchools?.(actor) ?? []);
  const greater = arrayIntersects(schools, spellFocusGreater?.getFocusedSchools?.(actor) ?? []);
  const mythic = arrayIntersects(schools, spellFocusMythic?.getFocusedSchools?.(actor) ?? []);
  let bonus = 0;
  if (focused) bonus += 1;
  if (greater) bonus += 1;
  if (mythic) bonus *= 2;
  return bonus;
}

function getElementalFocusFeatDcBonus(actor, actionUse) {
  if (!actor || actionUse?.item?.type !== "spell") return 0;
  const elements = getSpellElementKeys(actionUse);
  if (!elements.length) return 0;
  let hasFocus = false;
  let hasGreater = false;
  let hasMythic = false;
  for (const item of actor.items ?? []) {
    if (!isFeatSourceItem(item) || !isItemActive(item)) continue;
    const focused = (item.getFlag?.(CKL_MODULE_ID, "elemental-focus") ?? "").toString().trim().toLowerCase();
    const greater = (item.getFlag?.(CKL_MODULE_ID, "elemental-focus-greater") ?? "").toString().trim().toLowerCase();
    const mythic = (item.getFlag?.(CKL_MODULE_ID, "elemental-focus-mythic") ?? "").toString().trim().toLowerCase();
    if (focused && elements.includes(focused)) hasFocus = true;
    if (greater && elements.includes(greater)) hasGreater = true;
    if (mythic && elements.includes(mythic)) hasMythic = true;
  }
  let bonus = 0;
  if (hasFocus) bonus += 1;
  if (hasGreater) bonus += 1;
  if (hasMythic) bonus *= 2;
  return bonus;
}

function getElementalSpecificDcBonus(actor, actionUse) {
  if (!actor || actionUse?.item?.type !== "spell") return 0;
  const elements = getSpellElementKeys(actionUse);
  if (!elements.length) return 0;
  let total = 0;
  for (const item of actor.items ?? []) {
    if (!isFeatSourceItem(item) || !isItemActive(item)) continue;
    if (item?.hasItemBooleanFlag?.("elemental-dc") !== true) continue;
    const focusedElement = (item.getFlag?.(CKL_MODULE_ID, "elemental-dc") ?? "").toString().trim().toLowerCase();
    if (!focusedElement || !elements.includes(focusedElement)) continue;
    total += evaluateItemFormulaAsNumber(item, "elemental-dc-formula", actor);
  }
  return total;
}

function getElementalSpecificClBonus(actor, actionUse) {
  if (!actor || actionUse?.item?.type !== "spell") return 0;
  const elements = getSpellElementKeys(actionUse);
  if (!elements.length) return 0;
  let total = 0;
  for (const item of actor.items ?? []) {
    if (!isFeatSourceItem(item) || !isItemActive(item)) continue;
    if (item?.hasItemBooleanFlag?.("elemental-cl") !== true) continue;
    const focusedElement = (item.getFlag?.(CKL_MODULE_ID, "elemental-cl") ?? "").toString().trim().toLowerCase();
    if (!focusedElement || !elements.includes(focusedElement)) continue;
    total += evaluateItemFormulaAsNumber(item, "elemental-cl-formula", actor);
  }
  return total;
}

function getWeaponFocusAttackBonus(actor, actionUseLike) {
  if (!actor) return 0;
  const weaponFocus = getSpecificBonusType("weapon-focus");
  const weaponFocusGreater = getSpecificBonusType("weapon-focus-greater");
  const weaponFocusMythic = getSpecificBonusType("weapon-focus-mythic");
  const baseTypes = getActionBaseTypes(actionUseLike);
  if (!baseTypes.length) return 0;
  const focused = arrayIntersects(baseTypes, weaponFocus?.getFocusedWeapons?.(actor) ?? []);
  const greater = arrayIntersects(baseTypes, weaponFocusGreater?.getFocusedWeapons?.(actor) ?? []);
  const mythic = arrayIntersects(baseTypes, weaponFocusMythic?.getFocusedWeapons?.(actor) ?? []);
  let bonus = 0;
  if (focused) bonus += 1;
  if (greater) bonus += 1;
  if (mythic) bonus *= 2;
  return bonus;
}

function actionHasSaveDc(actionUse) {
  const action = actionUse?.action ?? actionUse?.shared?.action ?? null;
  if (!action) return false;
  if (action?.hasSave === true) return true;
  const saveType = (action?.save?.type ?? actionUse?.item?.system?.save?.type ?? "").toString().trim().toLowerCase();
  if (!saveType) return false;
  return saveType !== "none";
}

function isItemActive(item) {
  if (!item) return false;
  if (item.system?.disabled === true) return false;
  if (item.system?.active === false) return false;
  return true;
}

function onCklActionUseAlterRollData(actionUse) {
  const active = getActiveSpellPerfectionStateFromActionUse(actionUse);
  if (!active) return;
  if (!actionHasSaveDc(actionUse)) return;
  const shared = actionUse?.shared;
  if (!shared) return;
  if (shared.nasSpellPerfectionCklDcDoubled === true) return;
  const actor = actionUse?.actor ?? actionUse?.token?.actor ?? null;
  const featDcBonus =
    getSpellFocusFeatDcBonus(actor, actionUse) +
    getElementalFocusFeatDcBonus(actor, actionUse) +
    getTargetedFeatDcBonusFromUpdate(actionUse) +
    getElementalSpecificDcBonus(actor, actionUse);
  if (!Number.isFinite(featDcBonus) || featDcBonus === 0) return;
  shared.rollData.dcBonus = Number(shared?.rollData?.dcBonus ?? 0) + featDcBonus;
  shared.nasSpellPerfectionCklDcDoubled = true;
}

function onPreActorRollCl(actor, rollOptions) {
  if (!rollOptions?.messageId || !Array.isArray(rollOptions.parts)) return;
  const actionUse = game?.messages?.get?.(rollOptions.messageId)?.actionSource;
  if (!actionUse?.item || actionUse.item.type !== "spell") return;
  const active = getActiveSpellPerfectionStateFromActorAndAction(actor, actionUse);
  if (!active) return;
  const delta =
    getTargetedFeatClBonusFromUpdate(actionUse) +
    getElementalSpecificClBonus(actor, actionUse);
  if (!Number.isFinite(delta) || delta === 0) return;
  rollOptions.parts.push(`${delta}[Spell Perfection]`);
}

function onPreAttackRoll(action, _config, _rollData, _rollOptions, parts) {
  const actor = action?.actor;
  if (!actor) return;
  const active = getActiveSpellPerfectionStateFromActorAndAction(actor, action);
  if (!active) return;
  const bonus = getWeaponFocusAttackBonus(actor, action);
  if (!Number.isFinite(bonus) || bonus === 0) return;
  const originalParts = Array.isArray(parts) ? [...parts] : [];
  const nextParts = originalParts.filter((part) => {
    const value = (part ?? "").toString().trim();
    return !/\[.*weapon focus.*\]/i.test(value);
  });
  nextParts.push(`${bonus * 2}[Weapon Focus]`);
  parts.length = 0;
  parts.push(...nextParts);
}

export function registerSpellPerfection() {
  if (hooksRegistered) return;
  hooksRegistered = true;
  Hooks.on(CKL_ACTION_USE_ALTER_HOOK, onCklActionUseAlterRollData);
  Hooks.on("pf1PreActorRollCl", onPreActorRollCl);
  Hooks.on("pf1PreAttackRoll", onPreAttackRoll);
}
