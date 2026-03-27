import { MODULE } from "./module.js";

const fu = globalThis.foundry?.utils;
const duplicateFn = fu?.duplicate ?? globalThis.duplicate;
const mergeObjectFn = fu?.mergeObject ?? globalThis.mergeObject;

function entriesFromWorldStorage() {
  const storage = game.settings?.storage?.get?.("world");
  if (!storage) return [];
  const container = storage.contents ?? storage.settings ?? storage;
  if (container instanceof Map) {
    return Array.from(container.keys()).map(key => [key, container.get(key)]);
  }
  if (typeof container.entries === "function") {
    return Array.from(container.entries());
  }
  if (Array.isArray(container)) {
    return container.map((entry, index) => [entry?.key ?? index, entry]);
  }
  return [];
}

function worldStorageContainer() {
  const storage = game.settings?.storage?.get?.("world");
  if (!storage) return null;
  return storage.contents ?? storage.settings ?? storage;
}

function readWorldSettingValue(compositeKey) {
  const storage = game.settings?.storage?.get?.("world");
  if (!storage) return undefined;

  try {
    if (typeof storage.get === "function") {
      return storage.get(compositeKey)?.value;
    }
  } catch {
  }

  const container = worldStorageContainer();
  if (!container) return undefined;
  try {
    if (container instanceof Map) return container.get(compositeKey)?.value;
    if (typeof container.get === "function") return container.get(compositeKey)?.value;
    if (typeof container === "object") return container[compositeKey]?.value;
  } catch {
  }
  return undefined;
}

function writeWorldSettingValue(compositeKey, value) {
  const entry = { key: compositeKey, value };
  const storage = game.settings?.storage?.get?.("world");
  if (!storage) return false;

  try {
    if (typeof storage.set === "function") {
      storage.set(compositeKey, entry);
      return true;
    }
  } catch {
  }

  const container = worldStorageContainer();
  if (!container) return false;
  try {
    if (container instanceof Map) {
      container.set(compositeKey, entry);
      return true;
    }
    if (typeof container.set === "function") {
      container.set(compositeKey, entry);
      return true;
    }
    if (typeof container === "object") {
      container[compositeKey] = entry;
      return true;
    }
  } catch {
  }
  return false;
}

async function safeSetSetting(moduleId, key, value) {
  const compositeKey = `${moduleId}.${key}`;
  try {
    await game.settings.set(moduleId, key, value);
    return true;
  } catch (err) {
    const ok = writeWorldSettingValue(compositeKey, value);
    if (ok) {
      console.warn(`${MODULE.ID} | Migration | Wrote to world settings storage (fallback)`, { key: compositeKey, err });
    } else {
      console.warn(`${MODULE.ID} | Migration | Failed to write world setting (fallback)`, { key: compositeKey, err });
    }
    return ok;
  }
}

class MigrationProgress {
  constructor({ title, total }) {
    this.total = total;
    this.current = 0;
    this.detail = "";
    this._updateCount = 0;
    this.dialog = new Dialog({
      title,
      content: this.#renderContent(),
      buttons: {}
    }, {
      width: 520,
      height: "auto",
      resizable: false,
      classes: ["nas-migration-progress-dialog"]
    });
    this.dialog.render(true);

    setTimeout(() => {
      this.update({ current: this.current, detail: this.detail });
    }, 0);
  }

  update({ current, detail, status }) {
    if (typeof current === "number") this.current = current;
    if (detail !== undefined) this.detail = detail;
    const element = this.dialog?.element ?? this.dialog?._element;
    const root = element?.[0] ?? element?.get?.(0) ?? element;
    if (!root || typeof root.querySelector !== "function") return;
    const progressEl = root.querySelector("progress");
    const statusEl = root.querySelector("[data-migration-status]");
    const detailEl = root.querySelector("[data-migration-detail]");
    if (progressEl) {
      progressEl.value = this.current;
      progressEl.max = this.total;
    }
    if (statusEl) statusEl.textContent = status ?? `Migrating ${this.current} of ${this.total}`;
    if (detailEl) detailEl.textContent = this.detail || "";
  }

  async increment(detail) {
    this.update({ current: this.current + 1, detail });

    this._updateCount += 1;
    if (this._updateCount % 25 === 0) {
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
  }

  close() {
    this.dialog?.close();
  }

  #renderContent() {
    const detail = this.detail ? `<div data-migration-detail>${this.detail}</div>` : `<div data-migration-detail></div>`;
    return `
      <style>
        .nas-migration-progress-dialog .window-content {
          height: auto !important;
          overflow: visible !important;
        }
        .nas-migration-progress {
          display: flex;
          flex-direction: column;
          gap: 0.5rem;
        }
        .nas-migration-progress progress {
          width: 100%;
        }
      </style>
      <div class="nas-migration-progress">
        <p data-migration-status>Starting migration...</p>
        <progress value="${this.current}" max="${this.total}"></progress>
        ${detail}
      </div>
    `;
  }
}

function createMigrationProgress(total) {
  if (!game?.user?.isGM) return null;
  return new MigrationProgress({
    title: `${MODULE.NAME} Migration`,
    total
  });
}

export class MigrationSettingsMenu extends FormApplication {
  static get defaultOptions() {
    return mergeObjectFn(super.defaultOptions, {
      id: "nas-migration-menu",
      title: `${MODULE.NAME} Migration`,
      template: `modules/${MODULE.ID}/src/templates/migration-menu.html`,
      width: 460,
      height: "auto",
      resizable: false,
      closeOnSubmit: false
    });
  }

  getData() {
    return {
      moduleName: MODULE.NAME
    };
  }

  async _updateObject(event) {
    event?.preventDefault();
    await runSuiteMigrations({ force: true });
  }
}

async function migrateSettingsNamespace(legacyIds, newId) {
  const ids = Array.isArray(legacyIds) ? legacyIds : [legacyIds];
  const storage = game.settings?.storage?.get?.("world");
  if (!storage) return;

  const entries = entriesFromWorldStorage();

  for (const [rawKey, entry] of entries) {
    const key = typeof rawKey === "string" ? rawKey : entry?.key;
    if (!key || typeof key !== "string") continue;
    for (const legacyId of ids) {
      const prefix = `${legacyId}.`;
      if (!key.startsWith(prefix)) continue;
      const settingKey = key.slice(prefix.length);
      const newComposite = `${newId}.${settingKey}`;
      if (storage.get(newComposite)) continue;
      const legacyValue = storage.get(key)?.value;
      if (legacyValue === undefined) continue;
      await safeSetSetting(newId, settingKey, legacyValue);
    }
  }
}

function mergeLegacyFlags(doc, legacyIds, newId) {
  const ids = Array.isArray(legacyIds) ? legacyIds : [legacyIds];
  let merged = null;
  for (const legacyId of ids) {
    const legacyFlags = doc?.flags?.[legacyId];
    if (!legacyFlags || Object.keys(legacyFlags).length === 0) continue;
    const existing = merged ?? doc?.flags?.[newId] ?? {};
    merged = mergeObjectFn(duplicateFn(existing), legacyFlags, { inplace: false });
  }
  return merged;
}

async function migrateCollectionFlags(collection, legacyIds, newId, progress, labeler) {
  const ids = Array.isArray(legacyIds) ? legacyIds : [legacyIds];
  if (!collection) return;
  for (const doc of collection) {
    const merged = mergeLegacyFlags(doc, ids, newId);
    if (merged) {
      await doc.update({ [`flags.${newId}`]: merged });
    }
    if (progress) await progress.increment(labeler?.(doc) ?? doc.name ?? doc.id ?? "Document");
  }
}

async function migrateSceneTokens(legacyIds, newId, progress) {
  const ids = Array.isArray(legacyIds) ? legacyIds : [legacyIds];
  for (const scene of game.scenes ?? []) {
    const tokenUpdates = [];
    for (const token of scene.tokens ?? []) {
      const merged = mergeLegacyFlags(token, ids, newId);
      if (merged) tokenUpdates.push({ _id: token.id, [`flags.${newId}`]: merged });
      if (progress) await progress.increment(`Token: ${token.name ?? token.id} (Scene: ${scene.name ?? scene.id})`);
    }
    if (tokenUpdates.length > 0) {
      await scene.updateEmbeddedDocuments("Token", tokenUpdates);
    }
  }
}

async function migrateActorItems(legacyIds, newId, progress) {
  const ids = Array.isArray(legacyIds) ? legacyIds : [legacyIds];
  for (const actor of game.actors ?? []) {
    const itemUpdates = [];
    for (const item of actor.items ?? []) {
      const merged = mergeLegacyFlags(item, ids, newId);
      if (merged) {
        itemUpdates.push({ _id: item.id, [`flags.${newId}`]: merged });
      }
      if (progress) {
        await progress.increment(`Actor Item: ${item.name ?? item.id} (Actor: ${actor.name ?? actor.id})`);
      }
    }
    if (itemUpdates.length > 0) {
      await actor.updateEmbeddedDocuments("Item", itemUpdates);
    }
  }
}

async function migrateTokenActorItems(legacyIds, newId, progress) {
  const ids = Array.isArray(legacyIds) ? legacyIds : [legacyIds];
  for (const scene of game.scenes ?? []) {
    for (const token of scene.tokens ?? []) {
      const actor = token.actor;
      if (!actor) continue;
      const itemUpdates = [];
      for (const item of actor.items ?? []) {
        const merged = mergeLegacyFlags(item, ids, newId);
        if (merged) {
          itemUpdates.push({ _id: item.id, [`flags.${newId}`]: merged });
        }
        if (progress) {
          await progress.increment(`Token Item: ${item.name ?? item.id} (Token: ${token.name ?? token.id})`);
        }
      }
      if (itemUpdates.length > 0) {
        await actor.updateEmbeddedDocuments("Item", itemUpdates);
      }
    }
  }
}

async function migrateCustomDamageTypeFlags(legacyIds, newId) {
  const ids = Array.isArray(legacyIds) ? legacyIds : [legacyIds];
  let customDamageTypes;
  try {
    customDamageTypes = game.settings.get(newId, "customDamageTypes");
  } catch {
    customDamageTypes = readWorldSettingValue(`${newId}.customDamageTypes`);
  }
  if (!Array.isArray(customDamageTypes) || customDamageTypes.length === 0) return;
  let changed = false;
  for (const damageType of customDamageTypes) {
    const flags = damageType?.value?.flags;
    if (!flags) continue;
    for (const legacyId of ids) {
      if (!flags?.[legacyId]) continue;
      if (!flags[newId]) {
        flags[newId] = duplicateFn(flags[legacyId]);
      } else {
        flags[newId] = mergeObjectFn(duplicateFn(flags[newId]), flags[legacyId], { inplace: false });
      }
      changed = true;
    }
  }
  if (changed) {
    await safeSetSetting(newId, "customDamageTypes", customDamageTypes);
  }
}

function normalizeCustomDamageTypeEntry(entry, legacyId, newId) {
  if (!entry || typeof entry !== "object") return null;
  const rawKey = entry.key ?? entry.value?.name ?? "";
  const key = String(rawKey).toLowerCase().trim();
  if (!key) return null;

  const value = entry.value && typeof entry.value === "object" ? duplicateFn(entry.value) : {};
  value.name = value.name ?? entry.value?.name ?? rawKey;
  value.flags = value.flags && typeof value.flags === "object" ? value.flags : {};

  // Copy legacy module flags into the new module namespace (if present).
  const legacyFlags = value.flags?.[legacyId];
  if (legacyFlags && !value.flags[newId]) {
    value.flags[newId] = duplicateFn(legacyFlags);
  } else if (legacyFlags && value.flags[newId]) {
    value.flags[newId] = mergeObjectFn(duplicateFn(value.flags[newId]), legacyFlags, { inplace: false });
  }

  // Ensure NAS-style metadata exists (matches what `DamageTypeFormApplication` stores).
  value.namespace = newId;
  value._id = key;

  return { key, value };
}

async function migrateCustomDamageTypesFromLegacy(legacyId, newId) {
  let legacyTypes;
  try {
    legacyTypes = game.settings.get(legacyId, "customDamageTypes");
  } catch {
    legacyTypes = readWorldSettingValue(`${legacyId}.customDamageTypes`);
  }
  if (!Array.isArray(legacyTypes) || legacyTypes.length === 0) {
    return;
  }

  let currentTypes;
  try {
    currentTypes = game.settings.get(newId, "customDamageTypes");
  } catch {
    currentTypes = readWorldSettingValue(`${newId}.customDamageTypes`);
  }
  if (!Array.isArray(currentTypes)) currentTypes = [];

  const existingKeys = new Set(currentTypes.map((t) => String(t?.key ?? "").toLowerCase()));
  let imported = 0;
  let skipped = 0;

  for (const entry of legacyTypes) {
    const normalized = normalizeCustomDamageTypeEntry(entry, legacyId, newId);
    if (!normalized) continue;
    if (existingKeys.has(normalized.key)) { skipped += 1; continue; } // Do not overwrite existing NAS entries
    existingKeys.add(normalized.key);
    currentTypes.push(normalized);
    imported += 1;
  }

  if (imported > 0) {
    await safeSetSetting(newId, "customDamageTypes", currentTypes);
  }
}

function getRegisteredDefault(moduleId, key) {
  try {
    return game.settings?.settings?.get?.(`${moduleId}.${key}`)?.default;
  } catch {
    return undefined;
  }
}

function normalizeLegacyDamagePriorityName(typeName) {
  const type = String(typeName ?? "").trim();
  if (!type) return null;

  const materialTypes = globalThis.pf1?.registry?.materials;
  const alignments = globalThis.pf1?.config?.damageResistances;

  // Match materials by name/shortName; prefer treatedAs for canonical ids.
  const material = Array.isArray(materialTypes)
    ? materialTypes.find(m => m?.name === type || m?.shortName === type)
    : undefined;
  if (material) return material.treatedAs || material.id;

  if (alignments && typeof alignments === "object") {
    const alignmentKey = Object.keys(alignments).find(key => alignments[key] === type);
    if (alignmentKey) return alignmentKey;
  }

  return type.toLowerCase();
}

async function migrateDamageTypePriorityFromLegacy(legacyId, newId, { force } = {}) {
  let legacyPriority;
  try {
    legacyPriority = game.settings.get(legacyId, "damageTypePriority");
  } catch {
    legacyPriority = readWorldSettingValue(`${legacyId}.damageTypePriority`);
  }
  if (typeof legacyPriority === "string") {
    try {
      legacyPriority = JSON.parse(legacyPriority);
    } catch (e) {
      // Leave as-is; handled by array check below.
    }
  }
  if (!Array.isArray(legacyPriority) || legacyPriority.length === 0) {
    return;
  }

  let currentValue;
  try {
    currentValue = game.settings.get(newId, "damageTypePriority");
  } catch {
    currentValue = readWorldSettingValue(`${newId}.damageTypePriority`);
  }

  const defaultValue = getRegisteredDefault(newId, "damageTypePriority");
  const shouldOverwrite = !!force || !currentValue || currentValue === defaultValue;
  if (!shouldOverwrite) {
    return;
  }

  if (!globalThis.pf1?.registry?.materials || !globalThis.pf1?.config?.damageResistances) {
    console.warn(`${MODULE.ID} | Migration | damageTypePriority: PF1 registry/config not ready; mapping may be incomplete`);
  }

  const converted = legacyPriority.map(level => {
    if (!Array.isArray(level)) return [];
    return level
      .map(normalizeLegacyDamagePriorityName)
      .filter(Boolean);
  });

  await safeSetSetting(newId, "damageTypePriority", JSON.stringify(converted));
}

async function migrateTranslationsFromLegacy(legacyId, newId, { force } = {}) {
  let legacyTranslations;
  try {
    legacyTranslations = game.settings.get(legacyId, "translations");
  } catch {
    legacyTranslations = readWorldSettingValue(`${legacyId}.translations`);
  }
  if (!legacyTranslations || typeof legacyTranslations !== "object") {
    return;
  }

  let current;
  try {
    current = game.settings.get(newId, "translations");
  } catch {
    current = readWorldSettingValue(`${newId}.translations`);
  }
  if (!current || typeof current !== "object") current = {};

  const next = duplicateFn(current);
  let changed = false;
  const changedKeys = [];
  for (const k of ["hardness", "construct", "undead"]) {
    const incoming = legacyTranslations?.[k];
    if (typeof incoming !== "string") continue;
    const trimmed = incoming.trim();
    if (!trimmed) continue;

    const existing = typeof next[k] === "string" ? next[k].trim() : "";
    if (force || !existing) {
      next[k] = trimmed;
      changed = true;
      changedKeys.push(k);
    }
  }

  if (changed) {
    await safeSetSetting(newId, "translations", next);
  }
}

export async function runSuiteMigrations({ force = false } = {}) {
  const moduleVersion = game.modules?.get?.(MODULE.ID)?.version
    ?? game.modules?.get?.(MODULE.ID)?.data?.version
    ?? "0.0.0";
  const targetVersion = String(moduleVersion);
  let progress = null;
  let total = 0;
  try {
    // Only a GM can migrate world data and world-scope settings.
    if (!game?.user?.isGM) return;

    let storedVersion = "";
    try {
      storedVersion = game.settings.get(MODULE.ID, "migrationVersion") ?? "";
    } catch {
      storedVersion = readWorldSettingValue(`${MODULE.ID}.migrationVersion`) ?? "";
    }
    if (!force) {
      // If we've already migrated for this (or a newer) module version, skip.
      // Prefer Foundry's version comparison when available.
      try {
        const isNewer = globalThis.foundry?.utils?.isNewerVersion;
        if (typeof isNewer === "function") {
          if (!isNewer(targetVersion, storedVersion)) return;
        } else if (storedVersion === targetVersion) {
          return;
        }
      } catch {
        if (storedVersion === targetVersion) return;
      }
    }

    const legacyIds = [MODULE.LEGACY_AD, MODULE.LEGACY_IC];
    total = countMigrationTargets();
    progress = createMigrationProgress(total);

    await migrateSettingsNamespace(legacyIds, MODULE.ID);
    await migrateDamageTypePriorityFromLegacy(MODULE.LEGACY_AD, MODULE.ID, { force });
    await migrateTranslationsFromLegacy(MODULE.LEGACY_AD, MODULE.ID, { force });
    await migrateCustomDamageTypesFromLegacy(MODULE.LEGACY_AD, MODULE.ID);
    await migrateCustomDamageTypeFlags([MODULE.LEGACY_AD], MODULE.ID);
    await migrateCollectionFlags(game.actors, legacyIds, MODULE.ID, progress, actor => `Actor: ${actor.name ?? actor.id}`);
    await migrateActorItems(legacyIds, MODULE.ID, progress);
    await migrateCollectionFlags(game.items, legacyIds, MODULE.ID, progress, item => `World Item: ${item.name ?? item.id}`);
    await migrateCollectionFlags(game.combats, [MODULE.LEGACY_IC], MODULE.ID, progress, combat => `Combat: ${combat.name ?? combat.id}`);
    await migrateSceneTokens(legacyIds, MODULE.ID, progress);
    await migrateTokenActorItems(legacyIds, MODULE.ID, progress);

    progress?.update({ detail: "Finalizing migration..." });

    await safeSetSetting(MODULE.ID, "migrationVersion", targetVersion);
    progress?.update({
      current: total,
      status: "Migration complete",
      detail: "Migration complete."
    });
  } catch (err) {
    console.error(`${MODULE.ID} | Migration failed`, err);
    progress?.update({
      status: "Migration failed",
      detail: err?.message ? String(err.message) : "Migration failed. See console for details."
    });
  }
}

function countMigrationTargets() {
  let count = 0;
  count += game.actors?.size ?? 0;
  count += game.items?.size ?? 0;
  count += game.combats?.size ?? 0;
  for (const actor of game.actors ?? []) {
    count += actor.items?.size ?? 0;
  }
  for (const scene of game.scenes ?? []) {
    count += scene.tokens?.size ?? 0;
    for (const token of scene.tokens ?? []) {
      count += token.actor?.items?.size ?? 0;
    }
  }
  return count;
}



