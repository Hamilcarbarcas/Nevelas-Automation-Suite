import { MODULE } from "../../../common/module.js";

export function buildReactiveOptionChoices(optionList) {
  const sorted = [...optionList].sort((a, b) => a.label.localeCompare(b.label));
  const choices = {};
  const indexToId = [];
  for (let i = 0; i < sorted.length; i++) {
    choices[`c_${i}`] = sorted[i].label;
    indexToId.push(sorted[i].id);
  }
  return { choices, indexToId };
}

export function idsToTraitChoiceKeys(selectedIds, indexToId) {
  return selectedIds
    .map((id) => {
      const i = indexToId.indexOf(id);
      return i >= 0 ? `c_${i}` : null;
    })
    .filter(Boolean);
}

export function traitChoiceKeysToIds(keys, indexToId) {
  return keys
    .map((k) => {
      const m = /^c_(\d+)$/.exec(String(k));
      return m ? indexToId[Number(m[1])] : null;
    })
    .filter(Boolean);
}

export class ReactiveOptionSelector extends pf1.applications.ActorTraitSelector {
  constructor(options) {
    const {
      initialSelectedIds,
      indexToId,
      choices,
      onCommit,
      document: doc,
      title,
      subject,
      hasCustom,
      rowId,
      ...rest
    } = options;

    const dummySuffix = rowId ? String(rowId).replace(/[^a-zA-Z0-9_-]/g, "_") : foundry.utils.randomID();
    super({
      ...rest,
      document: doc,
      name: `flags.${MODULE.ID}._nasReactiveDummy_${dummySuffix}`,
      choices,
      subject: subject ?? "nasReactive",
      title: title ?? "",
      hasCustom: hasCustom ?? false,
    });

    this._nasIndexToId = indexToId;
    this._nasOnCommit = onCommit;

    this.attributes = { standard: new Set(), custom: new Set() };
    for (const k of idsToTraitChoiceKeys(initialSelectedIds ?? [], indexToId)) {
      if (choices[k]) this.attributes.standard.add(k);
    }
  }

  static async _updateDocumentNas(event, form, formData) {
    delete this.document.apps[this.appId];
    const { standard, custom } = this.attributes;
    const keys = [...standard.union(custom)];
    const indexToId = this._nasIndexToId;
    const selectedIds = traitChoiceKeysToIds(keys, indexToId);
    try {
      if (typeof this._nasOnCommit === "function") {
        await this._nasOnCommit(selectedIds);
      }
    } finally {
      this.close({ force: true });
    }
  }
}

const _baseDefaults = foundry.utils.deepClone(pf1.applications.ActorTraitSelector.DEFAULT_OPTIONS);
_baseDefaults.form = { ..._baseDefaults.form, handler: ReactiveOptionSelector._updateDocumentNas };
ReactiveOptionSelector.DEFAULT_OPTIONS = _baseDefaults;
