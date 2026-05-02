import { MODULE } from '../module.js';

export class VariantMappingManager extends FormApplication {
  static get defaultOptions() {
    return foundry.utils.mergeObject(super.defaultOptions, {
      id: 'variant-mapping-manager',
      title: game.i18n.localize("NAS.forms.variantMappingManager.title"),
      template: `modules/${MODULE.ID}/src/templates/variant-mapping-manager.html`,
      width: 600,
      height: 'auto',
      closeOnSubmit: true
    });
  }

  async getData() {
    const mappings = game.settings.get(MODULE.ID, 'pairedBuffMappings') || {};
    const rows = [];
    for (const [key, entry] of Object.entries(mappings)) {
      const displayName = await this._resolveSpellName(key);
      const allyLabel = await this._resolveBuffLabel(entry?.allies);
      const foeLabel = await this._resolveBuffLabel(entry?.foes);
      const switching = entry?.allowSwitching
      ? game.i18n.localize("NAS.common.buttons.yes")
      : game.i18n.localize("NAS.common.buttons.no");
      const perTarget = entry?.perTarget?.length ? `${entry.perTarget.length}` : '-';
      rows.push({ key, name: displayName, allies: allyLabel, foes: foeLabel, switching, perTarget });
    }
    rows.sort((a, b) => a.name.localeCompare(b.name));
    return { mappings: rows };
  }

  async _resolveSpellName(key) {
    try {
      if (key.startsWith('name:')) return key.substring(5);
      if (key.startsWith('item:')) {
        const id = key.substring(5);
        const item = game.items.get(id);
        return item?.name || key;
      }
      const doc = await fromUuid(key);
      return doc?.name || key;
    } catch (e) {
      return key;
    }
  }

  async _resolveBuffLabel(ref) {
    if (!ref) return '-';
    try {
      if (!ref.pack) {
        const item = game.items.get(ref.id);
        return item?.name || ref.id;
      }
      const pack = game.packs.get(ref.pack);
      if (!pack) return ref.id;
      const doc = await pack.getDocument(ref.id);
      return doc?.name || ref.id;
    } catch (e) {
      return ref.id;
    }
  }

  async _updateObject(event, formData) {
    const mappings = { ...(game.settings.get(MODULE.ID, 'pairedBuffMappings') || {}) };
    let changed = false;
    for (const [k, v] of Object.entries(formData)) {
      if (k.startsWith('del-') && v) {
        const key = k.substring(4);
        if (mappings[key] !== undefined) {
          delete mappings[key];
          changed = true;
        }
      }
    }
    if (changed) {
      await game.settings.set(MODULE.ID, 'pairedBuffMappings', mappings);
      ui.notifications.info(game.i18n.localize("NAS.forms.variantMappingManager.savedInfo"));
    }
  }
}
