import { MODULE } from '../module.js';
import { elementFromHtmlLike } from '../foundryCompat.js';

function getRegisteredDamageTypes() {
  const reg = pf1?.registry?.damageTypes;
  if (!reg?.entries) return [];
  const out = [];
  for (const [id, entry] of reg.entries()) {
    if (!id) continue;
    out.push({
      id: String(id),
      label: String(entry?.name ?? id)
    });
  }
  return out.sort((a, b) => a.label.localeCompare(b.label, undefined, { sensitivity: 'base' }));
}

export class WoundDamageTypesForm extends FormApplication {
  static get DEFAULT_TYPE_IDS() {
    return ["negative", "positive"];
  }

  static get defaultOptions() {
    return foundry.utils.mergeObject(super.defaultOptions, {
      id: "nas-wound-damage-types-form",
      title: game.i18n.localize("NAS.forms.woundDamageTypesForm.title"),
      template: "modules/nevelas-automation-suite/src/templates/wound-damage-types-form.html",
      width: 420,
      height: "auto",
      closeOnSubmit: true,
      classes: ["nas-wound-damage-types"]
    });
  }

  getData() {
    const selected = new Set(
      (game.settings.get(MODULE.ID, "woundsVigorWoundDamageTypeIds") ?? [])
        .map((id) => String(id).toLowerCase())
    );
    const types = getRegisteredDamageTypes().map((entry) => ({
      ...entry,
      checked: selected.has(entry.id.toLowerCase())
    }));
    return { types };
  }

  async _updateObject(_event, formData) {
    const selected = new Set();
    for (const [key, value] of Object.entries(formData)) {
      if (!key.startsWith("type.")) continue;
      if (!value) continue;
      const id = key.slice("type.".length).trim().toLowerCase();
      if (id) selected.add(id);
    }

    await game.settings.set(MODULE.ID, "woundsVigorWoundDamageTypeIds", Array.from(selected));
  }

  activateListeners(html) {
    super.activateListeners(html);

    const root = elementFromHtmlLike(html);
    if (!root) return;

    const setAll = (checked) => {
      root.querySelectorAll("input[type='checkbox'][name^='type.']").forEach((cb) => {
        cb.checked = checked;
      });
    };

    root.querySelector("[data-action='select-all']")?.addEventListener("click", () => setAll(true));
    root.querySelector("[data-action='clear-all']")?.addEventListener("click", () => setAll(false));
    root.querySelector("[data-action='defaults']")?.addEventListener("click", () => {
      const defaults = new Set(this.constructor.DEFAULT_TYPE_IDS.map((id) => String(id).toLowerCase()));
      root.querySelectorAll("input[type='checkbox'][name^='type.']").forEach((cb) => {
        const id = String(cb.name ?? "").slice("type.".length).toLowerCase();
        cb.checked = defaults.has(id);
      });
    });
  }
}
