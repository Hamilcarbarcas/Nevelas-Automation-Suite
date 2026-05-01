import { MODULE } from "../../../common/module.js";
import { elementFromHtmlLike } from "../../../common/foundryCompat.js";
import { buildReactiveOptionChoices, ReactiveOptionSelector } from "./reactiveOptionSelector.js";

const REACTIVE_FLAG_KEY = "itemReactiveEffects";

const ON_HIT_ACTION_SHEET_KEY = "onHitByActionOverride";
const TRACKED_APPS = new Map();
const SAVE_TIMEOUTS = new Map();
let BUFF_OPTIONS_CACHE = null;

class ReactiveUiState {
  static get(appId, key) {
    return TRACKED_APPS.get(appId)?.get(key);
  }

  static set(appId, key, value) {
    if (!TRACKED_APPS.has(appId)) TRACKED_APPS.set(appId, new Map());
    TRACKED_APPS.get(appId)?.set(key, value);

    const staleIds = [...TRACKED_APPS.keys()].filter((id) => !ui.windows?.[id]);
    for (const staleId of staleIds) TRACKED_APPS.delete(staleId);
  }
}

function localize(path) {
  return game.i18n.localize(`NAS.reactive.${path}`);
}

function isItemActionSheetContext(sheet) {
  return (
    sheet?.constructor?.name === "ItemActionSheet" ||
    (Array.isArray(sheet?.options?.classes) && sheet.options.classes.includes("item-action"))
  );
}

function deepClone(value) {
  return foundry.utils.deepClone(value ?? {});
}

function excludeNasChangeFromParentForm(event) {
  event?.stopPropagation?.();
}

function scheduleNasScrollRestoreRetry(sheet) {
  queueMicrotask(() => {
    if (!sheet || typeof sheet._restoreScrollPositions !== "function") return;
    const raw = sheet.element;
    const jq = raw?.jquery ? raw : typeof jQuery === "function" ? jQuery(raw?.[0] ?? raw) : null;
    if (!jq?.find) return;
    try {
      sheet._restoreScrollPositions(jq);
    } catch (_e) {
      return;
    }
  });
}

function syncReactiveSectionCollapsedChrome(section, expanded) {
  const body = section?.querySelector?.("[data-nas-reactive-body]");
  const header = section?.querySelector?.(".nas-reactive-section-header");
  if (body) body.style.display = expanded ? "" : "none";
  if (!header) return;
  if (expanded) {
    header.style.borderBottom = "";
  } else {
    header.style.borderBottom = "none";
  }
}

function findDetailsTab(root) {
  return root?.querySelector?.('.tab.details[data-group="primary"]') ?? null;
}

function findAdvancedHeaderInTab(tab) {
  if (!tab) return null;
  const advancedLabel = game.i18n.localize("PF1.Advanced").trim().toLowerCase();
  return (
    [...tab.querySelectorAll("h3.form-header")].find(
      (el) => (el.textContent ?? "").trim().toLowerCase() === advancedLabel
    ) ?? null
  );
}

function insertOnHitAtDetailsTabBottom(detailsTab, section) {
  const onStruck = detailsTab.querySelector(".nas-onstruck-effects");
  if (onStruck) {
    detailsTab.insertBefore(section, onStruck);
    return;
  }
  const advanced = findAdvancedHeaderInTab(detailsTab);
  if (advanced) {
    detailsTab.insertBefore(section, advanced);
    return;
  }
  detailsTab.appendChild(section);
}

function findLastActionTabSiblingInSection(fromHeader) {
  let n = fromHeader.nextElementSibling;
  let last = null;
  while (n) {
    if (n.matches?.("h3.form-header")) break;
    if (n.matches?.(".form-group, .form-groups") || n.classList?.contains?.("damage")) {
      last = n;
    }
    n = n.nextElementSibling;
  }
  return last;
}

function findActionTabDamageHealingFormHeader(hostTab) {
  const set = new Set();
  for (const key of ["PF1.DamageHealing", "PF1.DmgHealing", "PF1.DmgAndHealing", "PF1.ActionDamage", "PF1.Damage"]) {
    try {
      const t = game.i18n.localize(key)?.trim?.();
      if (t) set.add(t.toLowerCase());
    } catch {
    }
  }
  for (const h3 of hostTab.querySelectorAll("h3.form-header")) {
    const text = (h3.textContent ?? "").trim().toLowerCase();
    if (set.size && set.has(text)) return h3;
  }
  for (const h3 of hostTab.querySelectorAll("h3.form-header")) {
    const text = (h3.textContent ?? "").trim().toLowerCase();
    if (text.includes("damage") && (text.includes("heal") || text.includes("healing"))) return h3;
  }
  return null;
}

function insertOnHitInActionTab(hostTab, section) {
  const powerAttackInput = hostTab.querySelector('input[name="powerAttack.multiplier"]');
  const powerAttackGroup = powerAttackInput?.closest(".form-group") ?? null;
  const powerAttackHeader = powerAttackGroup?.previousElementSibling ?? null;
  if (powerAttackHeader?.matches?.("h3.form-header")) {
    hostTab.insertBefore(section, powerAttackHeader);
    return;
  }

  const damageBlocks = [...hostTab.querySelectorAll(".damage[data-key]")];
  if (damageBlocks.length > 0) {
    const lastBlock = damageBlocks[damageBlocks.length - 1];
    hostTab.insertBefore(section, lastBlock.nextSibling);
    return;
  }

  const damageHealingHeader = findActionTabDamageHealingFormHeader(hostTab);
  if (damageHealingHeader) {
    const lastInSection = findLastActionTabSiblingInSection(damageHealingHeader);
    const anchor = lastInSection ?? damageHealingHeader;
    hostTab.insertBefore(section, anchor.nextSibling);
    return;
  }

  hostTab.appendChild(section);
}

function getDamageTypeOptions() {
  const out = [{ id: "untyped", label: localize("damageTypeUntyped") }];
  for (const [, value] of pf1?.registry?.damageTypes?.entries?.() ?? []) {
    const id = String(value?.id ?? "").trim();
    if (!id) continue;
    out.push({ id, label: value?.name ?? id });
  }
  out.sort((a, b) => a.label.localeCompare(b.label));
  return out;
}

async function getBuffOptions() {
  if (Array.isArray(BUFF_OPTIONS_CACHE)) return BUFF_OPTIONS_CACHE;
  const selected = game.settings.get(MODULE.ID, "customBuffCompendia") || [];
  const includeWorld = selected.includes("__world__");
  const compendia = selected.filter((packId) => packId !== "__world__");
  const out = [];

  for (const packId of compendia) {
    const pack = game.packs.get(packId);
    if (!pack) continue;
    let index = [];
    try {
      index = await pack.getIndex();
    } catch (_err) {
      index = [];
    }
    const entries = index.filter((entry) => entry.type === "buff");
    for (const entry of entries) {
      out.push({
        id: `Compendium.${pack.collection}.${entry._id}`,
        label: entry.name
      });
    }
  }

  if (includeWorld) {
    for (const item of game.items ?? []) {
      if (item.type !== "buff") continue;
      out.push({ id: item.uuid, label: item.name });
    }
  }

  out.sort((a, b) => a.label.localeCompare(b.label));
  BUFF_OPTIONS_CACHE = out;
  return out;
}

function getConditionOptions() {
  const out = [];
  for (const condition of pf1?.registry?.conditions ?? []) {
    const id = String(condition?._id ?? "").trim();
    if (!id) continue;
    out.push({ id, label: condition?.name ?? id });
  }
  out.sort((a, b) => a.label.localeCompare(b.label));
  return out;
}

const ROW_ACTIONS = new Set(["applySelf", "removeSelf", "applyTarget", "removeTarget"]);

function normalizeRowAction(action) {
  const a = String(action ?? "");
  if (ROW_ACTIONS.has(a)) return a;
  if (a === "remove") return "removeSelf";
  return "applySelf";
}

function normalizeReactiveRows(rows = []) {
  if (!Array.isArray(rows)) return [];
  return rows
    .map((row) => ({
      id: String(row?.id ?? foundry.utils.randomID()),
      action: normalizeRowAction(row?.action),
      selectedIds: Array.isArray(row?.selectedIds) ? row.selectedIds.map((id) => String(id ?? "").trim()).filter(Boolean) : []
    }))
    .filter((row) => row.selectedIds.length > 0);
}

function persistReactiveRows(rows) {
  return normalizeReactiveRows(rows ?? []).map((r) => ({
    id: r.id,
    action: r.action,
    selectedIds: [...r.selectedIds]
  }));
}

function buffEffectToRowAction(effectType) {
  const t = String(effectType ?? "");
  if (t === "removeBuffAttacker") return "removeSelf";
  if (t === "applyBuffTarget") return "applyTarget";
  if (t === "removeBuffTarget") return "removeTarget";
  return "applySelf";
}

function conditionEffectToRowAction(effectType) {
  const t = String(effectType ?? "");
  if (t === "removeConditionAttacker") return "removeSelf";
  if (t === "applyConditionTarget") return "applyTarget";
  if (t === "removeConditionTarget") return "removeTarget";
  return "applySelf";
}

function buffRowsFromPersistedOrEffects(raw, effects) {
  if (raw != null && typeof raw === "object" && Object.prototype.hasOwnProperty.call(raw, "buffRows")) {
    return normalizeReactiveRows(raw.buffRows);
  }
  return effects
    .filter((effect) =>
      ["applyBuffAttacker", "removeBuffAttacker", "applyBuffTarget", "removeBuffTarget"].includes(String(effect?.type ?? "")) &&
      String(effect?.buffUuid ?? "")
    )
    .map((effect) => ({
      id: foundry.utils.randomID(),
      action: buffEffectToRowAction(effect?.type),
      selectedIds: [String(effect?.buffUuid ?? "")]
    }));
}

function conditionRowsFromPersistedOrEffects(raw, effects) {
  if (raw != null && typeof raw === "object" && Object.prototype.hasOwnProperty.call(raw, "conditionRows")) {
    return normalizeReactiveRows(raw.conditionRows);
  }
  return effects
    .filter((effect) =>
      ["applyConditionAttacker", "removeConditionAttacker", "applyConditionTarget", "removeConditionTarget"].includes(
        String(effect?.type ?? "")
      ) && String(effect?.conditionId ?? "")
    )
    .map((effect) => ({
      id: foundry.utils.randomID(),
      action: conditionEffectToRowAction(effect?.type),
      selectedIds: [String(effect?.conditionId ?? "")]
    }));
}

function attachReactiveRowEditor(section, item, state, rowKey, optionList, pickerTitle, onChange, reactiveContext = "onHit") {
  const list = section.querySelector(`[data-nas-list="${rowKey}"]`);
  const addBtn = section.querySelector(`[data-nas-add="${rowKey}"]`);
  if (!list || !addBtn) return;
  state[rowKey] = normalizeReactiveRows(state[rowKey]);
  const isOnStruck = reactiveContext === "onStruck";
  const optApplySelf = isOnStruck ? localize("actionOnStruckApplySelf") : localize("actionApplySelf");
  const optRemoveSelf = isOnStruck ? localize("actionOnStruckRemoveSelf") : localize("actionRemoveSelf");
  const optApplyAttacker = isOnStruck ? localize("actionOnStruckApplyAttacker") : localize("actionApplyTarget");
  const optRemoveAttacker = isOnStruck ? localize("actionOnStruckRemoveAttacker") : localize("actionRemoveTarget");

  const openTraitPicker = (row) => {
    if (!pf1?.applications?.ActorTraitSelector) {
      ui.notifications?.warn?.("PF1 trait selector is not available.");
      return;
    }
    const { choices, indexToId } = buildReactiveOptionChoices(optionList);
    const title = typeof pickerTitle === "string" ? pickerTitle : localize(String(pickerTitle ?? rowKey));
    new ReactiveOptionSelector({
      document: item,
      title,
      subject: `nasReactive-${rowKey}-${row.id}`,
      rowId: row.id,
      choices,
      indexToId,
      initialSelectedIds: [...row.selectedIds],
      hasCustom: false,
      onCommit: (selectedIds) => {
        row.selectedIds = selectedIds;
        render();
        onChange();
      },
    }).render(true);
  };

  const render = () => {
    list.innerHTML = "";
    for (const row of state[rowKey]) {
      const rowEl = document.createElement("div");
      rowEl.classList.add("nas-reactive-row");
      rowEl.dataset.rowId = row.id;
      rowEl.innerHTML = `
        <div class="nas-rx-fieldrow form-fields" style="display:flex;align-items:center;gap:6px;width:100%;margin:0;">
          <ul class="traits-list tag-list" data-role="tags" style="flex:1;min-width:0;min-height:var(--form-field-height,26px);margin:0;"></ul>
          <a data-role="edit" class="nas-reactive-trait-edit" title="${localize("editSelection")}" style="flex:0 0 auto;opacity:0.9;"><i class="fa-solid fa-edit" inert></i></a>
          <select data-role="action" style="flex:1 1 200px;min-width:min(100%,200px);max-width:280px;">
            <option value="applySelf">${optApplySelf}</option>
            <option value="removeSelf">${optRemoveSelf}</option>
            <option value="applyTarget">${optApplyAttacker}</option>
            <option value="removeTarget">${optRemoveAttacker}</option>
          </select>
          <a class="delete-row" data-role="remove" title="${localize("removeRow")}" style="flex:0 0 18px;text-align:center;"><i class="fas fa-trash"></i></a>
        </div>
      `;
      list.appendChild(rowEl);

      const tags = rowEl.querySelector('[data-role="tags"]');
      const editBtn = rowEl.querySelector('[data-role="edit"]');
      const actionSelect = rowEl.querySelector('[data-role="action"]');
      const removeBtn = rowEl.querySelector('[data-role="remove"]');

      const renderTags = () => {
        if (!row.selectedIds.length) {
          tags.innerHTML = `<li class="tag placeholder" inert>${game.i18n.localize("NAS.common.placeholders.noneSelected")}</li>`;
          return;
        }
        const labels = row.selectedIds
          .map((id) => optionList.find((option) => option.id === id)?.label ?? id)
          .sort((a, b) => a.localeCompare(b));
        tags.innerHTML = labels.map((label) => `<li class="tag">${foundry.utils.escapeHTML(label)}</li>`).join("");
      };

      editBtn.addEventListener("click", (event) => {
        event.preventDefault();
        event.stopPropagation();
        openTraitPicker(row);
      });
      actionSelect.value = row.action;
      actionSelect.addEventListener("change", (event) => {
        excludeNasChangeFromParentForm(event);
        row.action = normalizeRowAction(actionSelect.value);
        onChange();
      });
      removeBtn.addEventListener("click", () => {
        state[rowKey] = state[rowKey].filter((entry) => entry.id !== row.id);
        render();
        onChange();
      });

      renderTags();
    }
  };

  addBtn.addEventListener("click", () => {
    state[rowKey].push({
      id: foundry.utils.randomID(),
      action: "applySelf",
      selectedIds: []
    });
    render();
    onChange();
  });

  render();
}

function normalizeDamageTypeIds(ids) {
  if (!Array.isArray(ids) || !ids.length) return ["untyped"];
  const out = ids.map((id) => String(id ?? "").trim()).filter(Boolean);
  return out.length ? [...new Set(out)] : ["untyped"];
}

function attachDamageTypeMultiField(section, item, state, onChange) {
  const tags = section.querySelector("[data-nas-damage-type-tags]");
  const editBtn = section.querySelector("[data-nas-damage-type-edit]");
  if (!tags || !editBtn) return;

  state.damageTypeIds = normalizeDamageTypeIds(state.damageTypeIds);
  const dtSubject = section.classList.contains("nas-onhit-effects")
    ? "nasReactive-damageTypes-onhit"
    : "nasReactive-damageTypes-onstruck";

  const renderTags = () => {
    const opts = getDamageTypeOptions();
    const labels = state.damageTypeIds
      .map((id) => opts.find((o) => o.id === id)?.label ?? id)
      .sort((a, b) => a.localeCompare(b));
    tags.innerHTML =
      labels.length > 0
        ? labels.map((label) => `<li class="tag">${foundry.utils.escapeHTML(label)}</li>`).join("")
        : `<li class="tag placeholder" inert>${game.i18n.localize("NAS.common.placeholders.noneSelected")}</li>`;
  };

  const openPicker = () => {
    if (!pf1?.applications?.ActorTraitSelector) {
      ui.notifications?.warn?.("PF1 trait selector is not available.");
      return;
    }
    const optionList = getDamageTypeOptions();
    const { choices, indexToId } = buildReactiveOptionChoices(optionList);
    new ReactiveOptionSelector({
      document: item,
      title: game.i18n.localize("NAS.common.labels.damageTypes"),
      subject: dtSubject,
      rowId: "damageTypes",
      choices,
      indexToId,
      initialSelectedIds: [...state.damageTypeIds],
      hasCustom: false,
      onCommit: (selectedIds) => {
        state.damageTypeIds = normalizeDamageTypeIds(selectedIds);
        renderTags();
        onChange();
      },
    }).render(true);
  };

  editBtn.addEventListener("click", (event) => {
    event.preventDefault();
    event.stopPropagation();
    openPicker();
  });
  renderTags();
}

function primaryDamageHealFromEffects(effects) {
  return effects.find((effect) => ["healAttacker", "damageAttacker"].includes(String(effect?.type ?? "")));
}

function inferOnHitFunction(raw, effects) {
  const explicit = String(raw?.onHitFunction ?? "").trim();
  if (explicit === "lifesteal" || explicit === "none") return explicit;
  const preset = String(raw?.preset ?? "");
  if (preset === "lifesteal") return "lifesteal";
  const primary = primaryDamageHealFromEffects(effects);
  if (primary?.type === "healAttacker" && String(primary?.mode ?? "") === "percentOfFinalDamage") return "lifesteal";
  return "none";
}

function inferOnStruckFunction(raw, effects) {
  const explicit = String(raw?.onStruckFunction ?? "").trim();
  if (explicit === "none" || explicit === "damageAttacker" || explicit === "healAttacker") return explicit;
  const primary = primaryDamageHealFromEffects(effects);
  const preset = String(raw?.preset ?? "");
  if (preset === "none" && !primary) return "none";
  if ((preset === "custom" || preset === "") && !primary) return "none";
  if (preset === "fireShield" || preset === "thorns") return "damageAttacker";
  if (primary) return primary.type === "healAttacker" ? "healAttacker" : "damageAttacker";
  return "none";
}

function getReactiveFlags(item) {
  const raw = deepClone(item?.flags?.[MODULE.ID]?.[REACTIVE_FLAG_KEY] ?? {});
  raw.onHitByAction ??= {};
  raw[ON_HIT_ACTION_SHEET_KEY] ??= {};
  raw.onStruck ??= {};
  return raw;
}

function resolveReactivePostMessageFromRaw(raw, effects, primary) {
  if (raw != null && typeof raw === "object" && Object.prototype.hasOwnProperty.call(raw, "message") && typeof raw.message === "boolean") {
    return raw.message;
  }
  const list = Array.isArray(effects) ? effects : [];
  const withBool = list.find((e) => typeof e?.message === "boolean");
  if (typeof withBool?.message === "boolean") {
    return withBool.message;
  }
  return primary?.message !== false;
}

function normalizeOnHitConfig(raw = {}) {
  const effects = Array.isArray(raw?.effects) ? raw.effects : [];
  const primary = primaryDamageHealFromEffects(effects) ?? {};
  const message = resolveReactivePostMessageFromRaw(raw, effects, primary);
  const onHitFunction = inferOnHitFunction(raw, effects);
  const buffRows = buffRowsFromPersistedOrEffects(raw, effects);
  const conditionRows = conditionRowsFromPersistedOrEffects(raw, effects);
  const damageTypeIds = normalizeDamageTypeIds(
    Array.isArray(raw?.damageTypeIds) && raw.damageTypeIds.length
      ? raw.damageTypeIds
      : Array.isArray(primary?.damageTypes) && primary.damageTypes.length
        ? primary.damageTypes
        : primary?.damageType
          ? [String(primary.damageType)]
          : ["untyped"]
  );
  return {
    enabled: raw?.enabled === true,
    onHitFunction,
    mode: String(primary?.mode ?? "formula"),
    value: Number.isFinite(Number(primary?.value)) ? Number(primary.value) : 0,
    formula: String(primary?.formula ?? ""),
    damageTypeIds,
    buffRows,
    conditionRows,
    message
  };
}

function normalizeOnStruckConfig(raw = {}) {
  const effects = Array.isArray(raw?.effects) ? raw.effects : [];
  const primary = primaryDamageHealFromEffects(effects) ?? {};
  const message = resolveReactivePostMessageFromRaw(raw, effects, primary);
  const onStruckFunction = inferOnStruckFunction(raw, effects);
  const buffRows = buffRowsFromPersistedOrEffects(raw, effects);
  const conditionRows = conditionRowsFromPersistedOrEffects(raw, effects);
  const damageTypeIds = normalizeDamageTypeIds(
    Array.isArray(raw?.damageTypeIds) && raw.damageTypeIds.length
      ? raw.damageTypeIds
      : Array.isArray(primary?.damageTypes) && primary.damageTypes.length
        ? primary.damageTypes
        : primary?.damageType
          ? [String(primary.damageType)]
          : ["fire"]
  );
  return {
    enabled: raw?.enabled === true,
    onStruckFunction,
    mode: String(primary?.mode ?? "formula"),
    value: Number.isFinite(Number(primary?.value)) ? Number(primary.value) : 0,
    formula: String(primary?.formula ?? "1d6"),
    damageTypeIds,
    buffRows,
    conditionRows,
    message,
    meleeOnly: raw?.filters?.meleeOnly !== false,
    excludeReach: raw?.filters?.excludeReach !== false
  };
}

function mapBuffRowToEffectType(action, context = "onHit") {
  const a = normalizeRowAction(action);
  if (context === "onStruck") {
    if (a === "removeSelf") return "removeBuffTarget";
    if (a === "applyTarget") return "applyBuffAttacker";
    if (a === "removeTarget") return "removeBuffAttacker";
    return "applyBuffTarget";
  }
  if (a === "removeSelf") return "removeBuffAttacker";
  if (a === "applyTarget") return "applyBuffTarget";
  if (a === "removeTarget") return "removeBuffTarget";
  return "applyBuffAttacker";
}

function mapConditionRowToEffectType(action, context = "onHit") {
  const a = normalizeRowAction(action);
  if (context === "onStruck") {
    if (a === "removeSelf") return "removeConditionTarget";
    if (a === "applyTarget") return "applyConditionAttacker";
    if (a === "removeTarget") return "removeConditionAttacker";
    return "applyConditionTarget";
  }
  if (a === "removeSelf") return "removeConditionAttacker";
  if (a === "applyTarget") return "applyConditionTarget";
  if (a === "removeTarget") return "removeConditionTarget";
  return "applyConditionAttacker";
}

function toOnHitPayload(state) {
  const sectionMessage = state?.message !== false;
  if (!state?.enabled) {
    return {
      enabled: false,
      onHitFunction: "none",
      message: sectionMessage,
      effects: [],
      buffRows: [],
      conditionRows: []
    };
  }
  const effects = [];
  if (state.onHitFunction === "lifesteal") {
    const dt = normalizeDamageTypeIds(state.damageTypeIds);
    effects.push({
      type: "healAttacker",
      mode: String(state.mode ?? "percentOfFinalDamage"),
      value: Number(state.value) || 0,
      formula: String(state.formula ?? ""),
      damageTypes: dt,
      damageType: dt[0] ?? "untyped",
      message: state.message !== false
    });
  }
  for (const row of state.buffRows ?? []) {
    const effectType = mapBuffRowToEffectType(row?.action);
    for (const buffUuid of row?.selectedIds ?? []) {
      const uuid = String(buffUuid ?? "").trim();
      if (!uuid) continue;
      effects.push({ type: effectType, buffUuid: uuid, message: state.message !== false });
    }
  }
  for (const row of state.conditionRows ?? []) {
    const effectType = mapConditionRowToEffectType(row?.action);
    for (const conditionId of row?.selectedIds ?? []) {
      const id = String(conditionId ?? "").trim();
      if (!id) continue;
      effects.push({ type: effectType, conditionId: id, message: state.message !== false });
    }
  }
  return {
    enabled: true,
    onHitFunction: String(state.onHitFunction ?? "none"),
    message: sectionMessage,
    effects,
    buffRows: persistReactiveRows(state.buffRows),
    conditionRows: persistReactiveRows(state.conditionRows)
  };
}

function toOnStruckPayload(state) {
  const sectionMessage = state?.message !== false;
  if (!state?.enabled) {
    return {
      enabled: false,
      onStruckFunction: "none",
      message: sectionMessage,
      effects: [],
      buffRows: [],
      conditionRows: [],
      filters: {
        meleeOnly: state?.meleeOnly !== false,
        excludeReach: state?.excludeReach !== false
      }
    };
  }
  const effects = [];
  if (state.onStruckFunction === "damageAttacker" || state.onStruckFunction === "healAttacker") {
    const dt = normalizeDamageTypeIds(state.damageTypeIds);
    effects.push({
      type: String(state.onStruckFunction),
      mode: String(state.mode ?? "formula"),
      value: Number(state.value) || 0,
      formula: String(state.formula ?? ""),
      damageTypes: dt,
      damageType: dt[0] ?? "fire",
      message: state.message !== false
    });
  }
  for (const row of state.buffRows ?? []) {
    const effectType = mapBuffRowToEffectType(row?.action, "onStruck");
    for (const buffUuid of row?.selectedIds ?? []) {
      const uuid = String(buffUuid ?? "").trim();
      if (!uuid) continue;
      effects.push({ type: effectType, buffUuid: uuid, message: state.message !== false });
    }
  }
  for (const row of state.conditionRows ?? []) {
    const effectType = mapConditionRowToEffectType(row?.action, "onStruck");
    for (const conditionId of row?.selectedIds ?? []) {
      const id = String(conditionId ?? "").trim();
      if (!id) continue;
      effects.push({ type: effectType, conditionId: id, message: state.message !== false });
    }
  }
  return {
    enabled: true,
    onStruckFunction: String(state.onStruckFunction ?? "none"),
    message: sectionMessage,
    filters: {
      meleeOnly: state?.meleeOnly !== false,
      excludeReach: state?.excludeReach !== false
    },
    effects,
    buffRows: persistReactiveRows(state.buffRows),
    conditionRows: persistReactiveRows(state.conditionRows)
  };
}

function scheduleFlagSave(item, key, updater) {
  if (!item) return;
  const timeoutKey = `${item.uuid}:${key}`;
  const old = SAVE_TIMEOUTS.get(timeoutKey);
  if (old) clearTimeout(old);
  const nextTimer = setTimeout(async () => {
    SAVE_TIMEOUTS.delete(timeoutKey);
    const current = getReactiveFlags(item);
    const next = updater(current);
    await item.update({ [`flags.${MODULE.ID}.${REACTIVE_FLAG_KEY}`]: next }, { render: false });
  }, 300);
  SAVE_TIMEOUTS.set(timeoutKey, nextTimer);
}

function applyOnHitFunctionDefaults(state) {
  if (!state || state.onHitFunction !== "lifesteal") return;
  state.enabled = true;
  state.mode = "percentOfFinalDamage";
  state.value = 50;
  state.formula = "";
  state.damageTypeIds = normalizeDamageTypeIds(state.damageTypeIds?.length ? state.damageTypeIds : ["untyped"]);
  state.message = true;
}

async function renderOnHitSection(sheet, root) {
  const useActionSheetOverride = isItemActionSheetContext(sheet);
  const hostTab = useActionSheetOverride
    ? root.querySelector('.tab.action[data-group="primary"]')
    : findDetailsTab(root);
  if (!hostTab) return;
  const item = sheet?.item;
  const action = sheet?.action ?? [...(item?.actions ?? [])][0] ?? null;
  const shouldShowOnHit = useActionSheetOverride ? action?.hasDamage : item?.hasDamage;
  if (!shouldShowOnHit) {
    hostTab.querySelector(".nas-onhit-effects")?.remove();
    return;
  }
  const hadNas = !!hostTab.querySelector(".nas-onhit-effects");
  if (hadNas) return;
  if (!item || !action) return;
  const appKey = useActionSheetOverride ? "onhit-override" : "onhit";
  const fromState = ReactiveUiState.get(sheet.appId, appKey);
  const flags = getReactiveFlags(item);
  const rawOnHit = useActionSheetOverride
    ? flags[ON_HIT_ACTION_SHEET_KEY]?.[action.id]
    : flags.onHitByAction?.[action.id];
  const fromFlag = normalizeOnHitConfig(rawOnHit ?? {});
  const state = deepClone(fromState ?? fromFlag);
  state.buffRows = normalizeReactiveRows(state.buffRows);
  state.conditionRows = normalizeReactiveRows(state.conditionRows);
  if (!["none", "lifesteal"].includes(String(state.onHitFunction))) state.onHitFunction = fromFlag.onHitFunction;
  state.damageTypeIds = normalizeDamageTypeIds(
    Array.isArray(state.damageTypeIds) ? state.damageTypeIds : fromFlag.damageTypeIds
  );
  ReactiveUiState.set(sheet.appId, appKey, state);

  const section = document.createElement("div");
  section.classList.add("nas-onhit-effects");
  const onHitHeaderLabel = localize(useActionSheetOverride ? "onHitHeaderOverride" : "onHitHeader");
  section.innerHTML = `
    <h3 class="form-header nas-reactive-section-header" style="display:flex;flex-wrap:wrap;align-items:center;justify-content:space-between;gap:8px;">
      <span>${onHitHeaderLabel}</span>
      <label class="checkbox" style="margin:0;font-weight:normal;font-size:var(--font-size-14,0.875rem);" title="${localize("enabled")}">
        <input type="checkbox" data-nas-key="enabled" ${state.enabled ? "checked" : ""}>
      </label>
    </h3>
    <div data-nas-reactive-body>
    <div class="form-group nas-rx-function-row">
      <label class="nas-rx-function-label">${game.i18n.localize("NAS.reactive.labels.onHit")}</label>
      <div class="form-fields nas-rx-function-fields">
        <span class="nas-rx-arrow" aria-hidden="true">→</span>
        <select data-nas-key="onHitFunction">
          <option value="none">${game.i18n.localize("NAS.common.labels.none")}</option>
          <option value="lifesteal">${localize("presetLifesteal")}</option>
        </select>
      </div>
    </div>
    <div class="form-group" data-nas-row="mode">
      <label>${localize("mode")}</label>
      <div class="form-fields">
        <select data-nas-key="mode">
          <option value="percentOfFinalDamage">${localize("modePercentFinal")}</option>
          <option value="formula">${game.i18n.localize("NAS.common.labels.formula")}</option>
        </select>
      </div>
    </div>
    <div class="form-group" data-nas-row="value">
      <label>${localize("value")}</label>
      <div class="form-fields">
        <input type="number" step="1" data-nas-key="value" value="${Number(state.value) || 0}">
      </div>
    </div>
    <div class="form-group" data-nas-row="formula">
      <label>${game.i18n.localize("NAS.common.labels.formula")}</label>
      <div class="form-fields">
        <input class="formula roll" type="text" data-nas-key="formula" value="${state.formula ?? ""}" placeholder="${localize("formulaPlaceholder")}">
      </div>
    </div>
    <div class="form-group" data-nas-row="damageTypes">
      <label>${game.i18n.localize("NAS.common.labels.damageTypes")}</label>
      <div class="form-fields nas-reactive-dt-row" style="display:flex;align-items:center;gap:6px;width:100%;">
        <ul class="traits-list tag-list" data-nas-damage-type-tags style="flex:1;min-width:0;min-height:var(--form-field-height,26px);margin:0;"></ul>
        <a data-nas-damage-type-edit class="nas-reactive-trait-edit" title="${localize("editSelection")}" style="flex:0 0 auto;opacity:0.9;"><i class="fa-solid fa-edit" inert></i></a>
      </div>
    </div>
    <h4 class="form-header">${localize("additionalEffectsHeader")}</h4>
    <div class="nas-reactive-subheader" style="display:inline-flex;align-items:center;gap:8px;margin:4px 0 2px;flex-wrap:wrap;">
      <span style="font-weight:600;">${localize("buffsHeader")}</span>
      <a data-nas-add="buffRows" title="${localize("addRow")}" style="line-height:1;"><i class="fas fa-plus"></i></a>
    </div>
    <div data-nas-list="buffRows" style="margin:0 0 6px;"></div>
    <div class="nas-reactive-subheader" style="display:inline-flex;align-items:center;gap:8px;margin:2px 0 2px;flex-wrap:wrap;">
      <span style="font-weight:600;">${localize("conditionsHeader")}</span>
      <a data-nas-add="conditionRows" title="${localize("addRow")}" style="line-height:1;"><i class="fas fa-plus"></i></a>
    </div>
    <div data-nas-list="conditionRows" style="margin:0 0 6px;"></div>
    <div class="form-group stacked">
      <label class="checkbox">
        <input type="checkbox" data-nas-key="message" ${state.message ? "checked" : ""}>
        ${localize("postMessage")}
      </label>
    </div>
    </div>
  `;

  if (useActionSheetOverride) {
    insertOnHitInActionTab(hostTab, section);
  } else {
    insertOnHitAtDetailsTabBottom(hostTab, section);
  }

  const buffOptions = await getBuffOptions();
  const conditionOptions = getConditionOptions();
  const onHitFunctionSelect = section.querySelector('select[data-nas-key="onHitFunction"]');
  const modeSelect = section.querySelector('select[data-nas-key="mode"]');

  onHitFunctionSelect.value = String(state.onHitFunction ?? "none");
  modeSelect.value = String(state.mode ?? "formula");

  const updateRows = () => {
    syncReactiveSectionCollapsedChrome(section, state.enabled);

    const mode = state.mode;
    const fn = state.onHitFunction;
    const showPrimary = state.enabled && fn === "lifesteal";
    const showValue = showPrimary && mode === "percentOfFinalDamage";
    const showFormula = showPrimary && mode === "formula";
    const showDamageTypes = false;
    section.querySelector('[data-nas-row="value"]').style.display = showValue ? "" : "none";
    section.querySelector('[data-nas-row="formula"]').style.display = showFormula ? "" : "none";
    section.querySelector('[data-nas-row="damageTypes"]').style.display = showDamageTypes ? "" : "none";
    section.querySelector('[data-nas-row="mode"]').style.display = showPrimary ? "" : "none";
  };

  const writeState = () => {
    state.enabled = section.querySelector('input[data-nas-key="enabled"]')?.checked === true;
    state.onHitFunction = section.querySelector('select[data-nas-key="onHitFunction"]')?.value ?? "none";
    state.mode = section.querySelector('select[data-nas-key="mode"]')?.value ?? "formula";
    state.value = Number(section.querySelector('input[data-nas-key="value"]')?.value ?? 0) || 0;
    state.formula = String(section.querySelector('input[data-nas-key="formula"]')?.value ?? "");
    state.message = section.querySelector('input[data-nas-key="message"]')?.checked === true;
    ReactiveUiState.set(sheet.appId, appKey, state);
    updateRows();
    const saveDebounceKey = useActionSheetOverride ? `onhit-ov:${action.id}` : `onhit:${action.id}`;
    scheduleFlagSave(item, saveDebounceKey, (flags) => {
      if (useActionSheetOverride) {
        flags[ON_HIT_ACTION_SHEET_KEY] ??= {};
        flags[ON_HIT_ACTION_SHEET_KEY][action.id] = toOnHitPayload(state);
      } else {
        flags.onHitByAction ??= {};
        flags.onHitByAction[action.id] = toOnHitPayload(state);
      }
      return flags;
    });
  };

  attachDamageTypeMultiField(section, item, state, writeState);

  for (const control of section.querySelectorAll("input, select")) {
    control.addEventListener("change", (event) => {
      excludeNasChangeFromParentForm(event);
      if (control.dataset.nasKey === "onHitFunction") {
        state.onHitFunction = section.querySelector('select[data-nas-key="onHitFunction"]')?.value ?? "none";
        if (state.onHitFunction === "lifesteal") applyOnHitFunctionDefaults(state);
        section.querySelector('input[data-nas-key="enabled"]').checked = state.enabled === true;
        section.querySelector('select[data-nas-key="mode"]').value = state.mode;
        section.querySelector('input[data-nas-key="value"]').value = String(state.value ?? 0);
        section.querySelector('input[data-nas-key="formula"]').value = String(state.formula ?? "");
        section.querySelector('input[data-nas-key="message"]').checked = state.message === true;
      }
      writeState();
    });
  }

  const persistRows = () => {
    ReactiveUiState.set(sheet.appId, appKey, state);
    const saveDebounceKey = useActionSheetOverride ? `onhit-ov:${action.id}` : `onhit:${action.id}`;
    scheduleFlagSave(item, saveDebounceKey, (flags) => {
      if (useActionSheetOverride) {
        flags[ON_HIT_ACTION_SHEET_KEY] ??= {};
        flags[ON_HIT_ACTION_SHEET_KEY][action.id] = toOnHitPayload(state);
      } else {
        flags.onHitByAction ??= {};
        flags.onHitByAction[action.id] = toOnHitPayload(state);
      }
      return flags;
    });
  };
  attachReactiveRowEditor(section, item, state, "buffRows", buffOptions, localize("buffsHeader"), persistRows);
  attachReactiveRowEditor(section, item, state, "conditionRows", conditionOptions, localize("conditionsHeader"), persistRows);

  updateRows();
}

async function renderOnStruckSection(sheet, root) {
  const detailsTab = findDetailsTab(root);
  if (!detailsTab) return;
  const hadOnStruckNas = !!detailsTab.querySelector(".nas-onstruck-effects");
  const item = sheet?.item;
  if (!item || !["buff", "equipment"].includes(item.type)) return;
  if (hadOnStruckNas) return;

  const appKey = "onstruck";
  const fromState = ReactiveUiState.get(sheet.appId, appKey);
  const rawOnStruck = getReactiveFlags(item).onStruck;
  const fromFlag = normalizeOnStruckConfig(rawOnStruck);
  const state = deepClone(fromState ?? fromFlag);
  state.buffRows = normalizeReactiveRows(state.buffRows);
  state.conditionRows = normalizeReactiveRows(state.conditionRows);
  if (!["none", "damageAttacker", "healAttacker"].includes(String(state.onStruckFunction))) {
    state.onStruckFunction = fromFlag.onStruckFunction;
  }
  state.damageTypeIds = normalizeDamageTypeIds(
    Array.isArray(state.damageTypeIds) ? state.damageTypeIds : fromFlag.damageTypeIds
  );
  ReactiveUiState.set(sheet.appId, appKey, state);

  const section = document.createElement("div");
  section.classList.add("nas-onstruck-effects");
  section.innerHTML = `
    <h3 class="form-header nas-reactive-section-header" style="display:flex;flex-wrap:wrap;align-items:center;justify-content:space-between;gap:8px;">
      <span>${localize("onStruckHeader")}</span>
      <label class="checkbox" style="margin:0;font-weight:normal;font-size:var(--font-size-14,0.875rem);" title="${localize("enabled")}">
        <input type="checkbox" data-nas-key="enabled" ${state.enabled ? "checked" : ""}>
      </label>
    </h3>
    <div data-nas-reactive-body>
    <div class="form-group nas-rx-function-row">
      <label class="nas-rx-function-label">${game.i18n.localize("NAS.reactive.labels.onStruck")}</label>
      <div class="form-fields nas-rx-function-fields">
        <span class="nas-rx-arrow" aria-hidden="true">→</span>
        <select data-nas-key="onStruckFunction">
          <option value="none">${game.i18n.localize("NAS.common.labels.none")}</option>
          <option value="damageAttacker">${localize("effectDamageAttacker")}</option>
          <option value="healAttacker">${localize("effectHealAttacker")}</option>
        </select>
      </div>
    </div>
    <div class="form-group" data-nas-row="mode">
      <label>${localize("mode")}</label>
      <div class="form-fields">
        <select data-nas-key="mode">
          <option value="percentOfFinalDamage">${localize("modePercentFinal")}</option>
          <option value="formula">${game.i18n.localize("NAS.common.labels.formula")}</option>
        </select>
      </div>
    </div>
    <div class="form-group" data-nas-row="value">
      <label>${localize("value")}</label>
      <div class="form-fields">
        <input type="number" step="1" data-nas-key="value" value="${Number(state.value) || 0}">
      </div>
    </div>
    <div class="form-group" data-nas-row="formula">
      <label>${game.i18n.localize("NAS.common.labels.formula")}</label>
      <div class="form-fields">
        <input class="formula roll" type="text" data-nas-key="formula" value="${state.formula ?? ""}" placeholder="${localize("formulaPlaceholder")}">
      </div>
    </div>
    <div class="form-group" data-nas-row="damageTypes">
      <label>${game.i18n.localize("NAS.common.labels.damageTypes")}</label>
      <div class="form-fields nas-reactive-dt-row" style="display:flex;align-items:center;gap:6px;width:100%;">
        <ul class="traits-list tag-list" data-nas-damage-type-tags style="flex:1;min-width:0;min-height:var(--form-field-height,26px);margin:0;"></ul>
        <a data-nas-damage-type-edit class="nas-reactive-trait-edit" title="${localize("editSelection")}" style="flex:0 0 auto;opacity:0.9;"><i class="fa-solid fa-edit" inert></i></a>
      </div>
    </div>
    <h4 class="form-header">${localize("additionalEffectsHeader")}</h4>
    <div class="nas-reactive-subheader" style="display:inline-flex;align-items:center;gap:8px;margin:4px 0 2px;flex-wrap:wrap;">
      <span style="font-weight:600;">${localize("buffsHeader")}</span>
      <a data-nas-add="buffRows" title="${localize("addRow")}" style="line-height:1;"><i class="fas fa-plus"></i></a>
    </div>
    <div data-nas-list="buffRows" style="margin:0 0 6px;"></div>
    <div class="nas-reactive-subheader" style="display:inline-flex;align-items:center;gap:8px;margin:2px 0 2px;flex-wrap:wrap;">
      <span style="font-weight:600;">${localize("conditionsHeader")}</span>
      <a data-nas-add="conditionRows" title="${localize("addRow")}" style="line-height:1;"><i class="fas fa-plus"></i></a>
    </div>
    <div data-nas-list="conditionRows" style="margin:0 0 6px;"></div>
    <div class="form-group stacked">
      <label class="checkbox">
        <input type="checkbox" data-nas-key="meleeOnly" ${state.meleeOnly ? "checked" : ""}>
        ${localize("meleeOnly")}
      </label>
      <label class="checkbox">
        <input type="checkbox" data-nas-key="excludeReach" ${state.excludeReach ? "checked" : ""}>
        ${localize("excludeReach")}
      </label>
      <label class="checkbox">
        <input type="checkbox" data-nas-key="message" ${state.message ? "checked" : ""}>
        ${localize("postMessage")}
      </label>
    </div>
    </div>
  `;
  const advancedHeader = findAdvancedHeaderInTab(detailsTab);
  if (advancedHeader) {
    detailsTab.insertBefore(section, advancedHeader);
  } else {
    const firstAnchor = detailsTab.querySelector("h3.form-header, .form-group, hr");
    if (firstAnchor) detailsTab.insertBefore(section, firstAnchor);
    else detailsTab.appendChild(section);
  }

  const buffOptions = await getBuffOptions();
  const conditionOptions = getConditionOptions();
  const onStruckFunctionSelect = section.querySelector('select[data-nas-key="onStruckFunction"]');
  const modeSelect = section.querySelector('select[data-nas-key="mode"]');
  onStruckFunctionSelect.value = String(state.onStruckFunction ?? "none");
  modeSelect.value = String(state.mode ?? "formula");

  const updateRows = () => {
    syncReactiveSectionCollapsedChrome(section, state.enabled);

    const fn = state.onStruckFunction;
    const damageLike = fn === "damageAttacker" || fn === "healAttacker";
    const showValue = state.enabled && damageLike && state.mode === "percentOfFinalDamage";
    const showFormula = state.enabled && damageLike && state.mode === "formula";
    const showMode = state.enabled && damageLike;
    const showDamageTypes = state.enabled && fn === "damageAttacker";
    section.querySelector('[data-nas-row="mode"]').style.display = showMode ? "" : "none";
    section.querySelector('[data-nas-row="value"]').style.display = showValue ? "" : "none";
    section.querySelector('[data-nas-row="formula"]').style.display = showFormula ? "" : "none";
    section.querySelector('[data-nas-row="damageTypes"]').style.display = showDamageTypes ? "" : "none";
  };

  const writeState = () => {
    state.enabled = section.querySelector('input[data-nas-key="enabled"]')?.checked === true;
    state.onStruckFunction = section.querySelector('select[data-nas-key="onStruckFunction"]')?.value ?? "none";
    state.mode = section.querySelector('select[data-nas-key="mode"]')?.value ?? "formula";
    state.value = Number(section.querySelector('input[data-nas-key="value"]')?.value ?? 0) || 0;
    state.formula = String(section.querySelector('input[data-nas-key="formula"]')?.value ?? "");
    state.message = section.querySelector('input[data-nas-key="message"]')?.checked === true;
    state.meleeOnly = section.querySelector('input[data-nas-key="meleeOnly"]')?.checked === true;
    state.excludeReach = section.querySelector('input[data-nas-key="excludeReach"]')?.checked === true;
    ReactiveUiState.set(sheet.appId, appKey, state);
    updateRows();
    scheduleFlagSave(item, "onstruck", (flags) => {
      flags.onStruck = toOnStruckPayload(state);
      return flags;
    });
  };

  attachDamageTypeMultiField(section, item, state, writeState);

  for (const control of section.querySelectorAll("input, select")) {
    control.addEventListener("change", (event) => {
      excludeNasChangeFromParentForm(event);
      if (control.dataset.nasKey === "onStruckFunction") {
        state.onStruckFunction = section.querySelector('select[data-nas-key="onStruckFunction"]')?.value ?? "none";
        section.querySelector('input[data-nas-key="enabled"]').checked = state.enabled === true;
        section.querySelector('select[data-nas-key="mode"]').value = state.mode;
        section.querySelector('input[data-nas-key="value"]').value = String(state.value ?? 0);
        section.querySelector('input[data-nas-key="formula"]').value = String(state.formula ?? "");
        section.querySelector('input[data-nas-key="message"]').checked = state.message === true;
        section.querySelector('input[data-nas-key="meleeOnly"]').checked = state.meleeOnly === true;
        section.querySelector('input[data-nas-key="excludeReach"]').checked = state.excludeReach === true;
      }
      writeState();
    });
  }

  const persistRows = () => {
    ReactiveUiState.set(sheet.appId, appKey, state);
    scheduleFlagSave(item, "onstruck", (flags) => {
      flags.onStruck = toOnStruckPayload(state);
      return flags;
    });
  };
  attachReactiveRowEditor(section, item, state, "buffRows", buffOptions, localize("buffsHeader"), persistRows, "onStruck");
  attachReactiveRowEditor(section, item, state, "conditionRows", conditionOptions, localize("conditionsHeader"), persistRows, "onStruck");

  updateRows();
}

function onRenderItemActionSheet(sheet, html) {
  const root = elementFromHtmlLike(sheet?.element) ?? elementFromHtmlLike(html);
  if (!root) return;
  void renderOnHitSection(sheet, root).finally(() => scheduleNasScrollRestoreRetry(sheet));
}

function onRenderItemSheet(sheet, html) {
  const root = elementFromHtmlLike(sheet?.element) ?? elementFromHtmlLike(html);
  if (!root) return;
  void renderOnHitSection(sheet, root)
    .then(() => renderOnStruckSection(sheet, root))
    .finally(() => scheduleNasScrollRestoreRetry(sheet));
}

export function registerReactiveItemSheet() {
  if (game?.ready) void getBuffOptions();
  else Hooks.once("ready", () => void getBuffOptions());
  Hooks.on("renderItemActionSheet", onRenderItemActionSheet);
  Hooks.on("renderItemSheetPF", onRenderItemSheet);
}
