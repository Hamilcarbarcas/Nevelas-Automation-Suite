import { resolveEnglishName } from "./compendiumNameResolver.js";
import { resolveMetamagicNameFromDatabase } from "../metamagic/metamagic.js";

export const GRAPPLE_FORM_KEY = "grapple";
export const METAMAGIC_FORM_KEY = "metamagic";
export const METAMAGIC_SELECT_KEY = "metamagicSelection";
export const METAMAGIC_NAMES_KEY = "metamagicNames";
export const METAMAGIC_DROPDOWN_KEY = "metamagicDropdownOpen";
export const METAMAGIC_OPTIONS_KEY = "metamagicOptions";

class DialogStateTracker {
  static #trackedApplications = new Map();

  static get(appId, key) {
    return this.#trackedApplications.get(appId)?.get(key);
  }

  static set(appId, key, value) {
    if (!this.#trackedApplications.has(appId)) {
      this.#trackedApplications.set(appId, new Map());
    }
    this.#trackedApplications.get(appId)?.set(key, value);

    const toRemove = [...this.#trackedApplications.keys()].filter((id) => !ui.windows?.[id]);
    for (const id of toRemove) {
      this.#trackedApplications.delete(id);
    }
  }
}

function getFlagsContainer(form) {
  let container = form.querySelector('div.form-group.stacked.flags');
  if (container) return container;

  container = document.createElement('div');
  container.classList.add('form-group', 'stacked', 'flags');
  const label = document.createElement('label');
  label.textContent = ` ${game.i18n.localize('PF1.Misc')} `;
  container.appendChild(label);
  const sibling = form.querySelector('.form-group.flags') || form.querySelector('.form-group');
  if (sibling) sibling.after(container);
  else form.appendChild(container);
  return container;
}

function getActionItemType(dialog) {
  return dialog?.action?.itemType ?? dialog?.action?.item?.type ?? dialog?.action?.item?.data?.type;
}

function getActionActor(dialog) {
  return dialog?.action?.actor ?? dialog?.action?.item?.actor ?? dialog?.actor ?? null;
}

function getSpellComponents(dialog) {
  return (
    dialog?.action?.components ??
    dialog?.action?.item?.system?.components ??
    dialog?.action?.item?.data?.components ??
    {}
  );
}

function getSpellDuration(dialog) {
  return dialog?.action?.duration ?? dialog?.action?.item?.system?.duration ?? {};
}

function getSpellActivation(dialog) {
  return dialog?.action?.activation ?? dialog?.action?.item?.system?.activation ?? {};
}

function getSpellDamageParts(dialog) {
  return (
    dialog?.action?.damage?.parts ??
    dialog?.action?.item?.system?.damage?.parts ??
    []
  );
}

function hasDamageFormula(dialog) {
  const parts = getSpellDamageParts(dialog);
  if (!Array.isArray(parts) || !parts.length) return false;
  return parts.some((part) => {
    const formula = part?.formula ?? part?.[0];
    return typeof formula === "string" && formula.trim().length > 0;
  });
}

function getSpellCasterLevel(dialog) {
  const cl = dialog?.rollData?.cl;
  const value = Number(cl ?? 0);
  return Number.isFinite(value) ? value : 0;
}

function getSpellSaveType(dialog) {
  return (
    dialog?.action?.save?.type ??
    dialog?.action?.item?.system?.save?.type ??
    ""
  );
}

function getSpellRangeUnits(dialog) {
  const rawUnits =
    dialog?.action?.range?.units ??
    dialog?.action?.item?.system?.range?.units ??
    dialog?.action?.item?.data?.range?.units ??
    "";
  const normalized = rawUnits?.toString?.().toLowerCase() ?? "";
  if (normalized) return normalized;
  if (dialog?.action?.touch) return "touch";
  return "";
}

function getCanonicalMetamagicName(source) {
  const raw = source?.metaName ?? source?.label ?? "";
  return resolveMetamagicNameFromDatabase(raw) ?? raw;
}

function canApplyExtendSpell(duration) {
  const units = (duration?.units ?? "").toString().toLowerCase();
  if (duration?.concentration) return false;
  if (units === "inst" || units === "instantaneous" || units === "perm" || units === "permanent") {
    return false;
  }
  return true;
}

function isAreaEffectSpell(dialog) {
  const areaString = dialog?.action?.area ?? dialog?.action?.item?.system?.area ?? "";
  const templateType = dialog?.action?.measureTemplate?.type ?? dialog?.action?.item?.system?.measureTemplate?.type;
  return Boolean(areaString || templateType);
}

function isInstantDuration(duration) {
  const units = (duration?.units ?? "").toString().toLowerCase();
  return units === "inst" || units === "instantaneous";
}

function getSpellbookAbilityMod(dialog) {
  const mod = dialog?.rollData?.ablMod;
  const value = Number(mod ?? 0);
  return Number.isFinite(value) ? value : 0;
}

function getSpellBaseLevel(dialog) {
  const level = dialog?.action?.item?.system?.level ?? dialog?.item?.system?.level;
  const value = Number(level ?? 0);
  return Number.isFinite(value) ? value : 0;
}

function canApplyQuickenSpell(activation) {
  const type = (activation?.type ?? "").toString().toLowerCase();
  const costValue = Number(activation?.cost ?? 1);
  const cost = Number.isFinite(costValue) ? costValue : 1;
  if (type === "swift") return false;
  if (!type) return true;
  if (["round", "full"].includes(type)) {
    return cost <= 1;
  }
  if (["minute", "hour", "day", "week", "month", "year"].includes(type)) return false;
  return true;
}

function canApplyIntensifiedSpell(dialog) {
  const cl = getSpellCasterLevel(dialog);
  if (!Number.isFinite(cl) || cl <= 0) return false;
  const parts = getSpellDamageParts(dialog);
  if (!Array.isArray(parts) || !parts.length) return false;
  const half = Math.floor(cl / 2);
  const doubled = cl * 2;

  const hasCap = (formula) => {
    if (!formula || typeof formula !== "string") return false;
    const dicePattern = /(\([^)]*\)|@cl|clamp\([^)]*\)|min\([^)]*\)|floor\([^)]*\)|\d+)\s*\)*\s*d\s*\d+/i;
    if (!dicePattern.test(formula)) {
      return false;
    }

    const clampMatch = formula.match(/clamp\(\s*floor\(\s*@cl\s*\/\s*2\s*\)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)/i);
    if (clampMatch) {
      const cap = Number(clampMatch[2]);
      return Number.isFinite(cap) && half > cap;
    }

    const minCapHalf = formula.match(/min\(\s*(\d+)\s*,\s*floor\(\s*@cl\s*\/\s*2\s*\)\s*\)/i);
    if (minCapHalf) {
      const cap = Number(minCapHalf[1]);
      return Number.isFinite(cap) && half > cap;
    }

    const minHalfCap = formula.match(/min\(\s*floor\(\s*@cl\s*\/\s*2\s*\)\s*,\s*(\d+)\s*\)/i);
    if (minHalfCap) {
      const cap = Number(minHalfCap[1]);
      return Number.isFinite(cap) && half > cap;
    }

    const minCapCl = formula.match(/min\(\s*(\d+)\s*,\s*@cl\s*\)/i);
    if (minCapCl) {
      const cap = Number(minCapCl[1]);
      return Number.isFinite(cap) && cl > cap;
    }

    const minClCap = formula.match(/min\(\s*@cl\s*,\s*(\d+)\s*\)/i);
    if (minClCap) {
      const cap = Number(minClCap[1]);
      return Number.isFinite(cap) && cl > cap;
    }

    const minCapDouble = formula.match(/min\(\s*(\d+)\s*,\s*@cl\s*\*\s*2\s*\)/i);
    if (minCapDouble) {
      const cap = Number(minCapDouble[1]);
      return Number.isFinite(cap) && doubled > cap;
    }

    return false;
  };

  return parts.some((part) => {
    const formula = part?.formula ?? part?.[0];
    return hasCap(formula);
  });
}

function filterMetamagicSourcesForDialog(dialog, sources) {
  const components = getSpellComponents(dialog);
  const duration = getSpellDuration(dialog);
  const activation = getSpellActivation(dialog);
  const rangeUnits = getSpellRangeUnits(dialog);
  const isAreaEffect = isAreaEffectSpell(dialog);
  const abilityMod = getSpellbookAbilityMod(dialog);

  return sources.filter((source) => {
    const name = getCanonicalMetamagicName(source).toString().toLowerCase();
    if (name === "still spell") {
      return components?.somatic === true;
    }
    if (name === "silent spell") {
      return components?.verbal === true;
    }
    if (name === "extend spell") {
      return canApplyExtendSpell(duration);
    }
    if (name === "reach spell") {
      return ["touch", "close", "medium"].includes(rangeUnits);
    }
    if (name === "quicken spell") {
      return canApplyQuickenSpell(activation);
    }
    if (name === "selective spell") {
      return isAreaEffect && isInstantDuration(duration) && abilityMod > 0;
    }
    if (name === "dazing spell") {
      return hasDamageFormula(dialog);
    }
    if (name === "persistent spell") {
      return Boolean(getSpellSaveType(dialog));
    }
    if (name === "heighten spell") {
      const baseLevel = getSpellBaseLevel(dialog);
      return baseLevel < 9;
    }
    if (name === "intensified spell") {
      return canApplyIntensifiedSpell(dialog);
    }
    if (name === "maximize spell") {
      return hasDamageFormula(dialog);
    }
    return true;
  });
}

function hasAvailableDailyUses(item) {
  const uses = item?.system?.uses;
  if (!uses) return true;
  if (uses?.per !== "day") return true;
  const value = Number(uses?.value ?? 0);
  return Number.isFinite(value) && value > 0;
}

function isMetamagicFeat(item) {
  if (item?.type !== "feat" || item?.subType !== "feat") return false;
  const tags = item?.system?.tags;
  if (!Array.isArray(tags)) return false;
  return tags.some((tag) => tag?.toString?.().toLowerCase().includes("metamagic"));
}

function isMetamagicRodName(name) {
  return Boolean(extractRodMetamagicPrefix(name));
}

function cleanMetamagicName(name) {
  return (name ?? "")
    .toString()
    .replace(/[,]/g, " ")
    .replace(/\b(lesser|greater)\b/gi, " ")
    .replace(/\bmetamagic\b/gi, " ")
    .replace(/\brod\b/gi, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function extractRodMetamagicPrefix(name) {
  const label = (name ?? "").toString();
  if (!label) return "";
  const lower = label.toLowerCase();
  if (!lower.includes("rod") || !lower.includes("metamagic")) return "";

  const ofMatch = label.match(/rod\s+of\s+(.+?)\s+metamagic/i);
  if (ofMatch?.[1]) return cleanMetamagicName(ofMatch[1]);

  const rodMatch = label.match(/(.+?)\s+metamagic\s+rod/i);
  if (rodMatch?.[1]) return cleanMetamagicName(rodMatch[1]);

  const suffixMatch = label.match(/metamagic\s+rod[, ]+\s*(.+)$/i);
  if (suffixMatch?.[1]) return cleanMetamagicName(suffixMatch[1]);

  return "";
}

async function getItemEnglishName(item) {
  const itemName = (item?.name ?? "").toString().trim();
  if (!itemName) return itemName;

  const databaseName = resolveMetamagicNameFromDatabase(itemName);
  if (databaseName) return databaseName;

  return resolveEnglishName(itemName, { documentName: "Item", deepScanMode: "off" });
}

async function getAvailableMetamagicSources(actor, options = {}) {
  if (!actor?.items?.size) return [];
  const { resolveEnglishNames = false } = options;
  const items = Array.from(actor.items).filter(
    (item) =>
      (item?.type === "feat" && item?.subType === "feat") ||
      item?.constructor?.name === "ItemEquipmentPF"
  );
  items.sort((a, b) => {
    const aHasOriginal = Boolean(a?.flags?.babele?.originalName);
    const bHasOriginal = Boolean(b?.flags?.babele?.originalName);
    return Number(bHasOriginal) - Number(aHasOriginal);
  });
  const resolved = await Promise.all(
    items.map(async (item) => {
      const originalName = item?.flags?.babele?.originalName;
      const englishName = originalName
        ? originalName
        : resolveEnglishNames
          ? await getItemEnglishName(item)
          : item?.name ?? "";

      return { item, englishName };
    })
  );

  return resolved
    .map(({ item, englishName }) => {
      if (!item || !englishName) return null;

      if (isMetamagicFeat(item)) {
        return {
          item,
          type: "feat",
          label: item?.name ?? cleanMetamagicName(englishName),
          metaName: englishName,
        };
      }

      const rodPrefix =
        extractRodMetamagicPrefix(englishName) ||
        extractRodMetamagicPrefix(item?.flags?.babele?.originalName) ||
        extractRodMetamagicPrefix(item?.name);
      if (!rodPrefix) return null;
      if (!hasAvailableDailyUses(item)) return null;

      return {
        item,
        type: "rod",
        label: item?.name ?? rodPrefix,
        metaName: englishName,
      };
    })
    .filter(Boolean);
}

export function addGrappleCheckbox(dialog, html) {
  if (!(dialog instanceof pf1.applications.AttackDialog)) return;
  if (!dialog.action?.hasAttack) return;

  const root = Array.isArray(html) ? html[0] : html?.[0] || html;
  if (!root) return;
  const form = root.querySelector?.('form') ?? root;
  if (!form) return;

  if (form.querySelector(`input[name="${GRAPPLE_FORM_KEY}"]`)) return;

  const labelText = game.i18n.localize('NAS.conditions.main.GrappleCheckbox');
  const container = getFlagsContainer(form);

  const labelElement = document.createElement('label');
  labelElement.classList.add('checkbox');
  const input = document.createElement('input');
  input.type = 'checkbox';
  input.name = GRAPPLE_FORM_KEY;
  const storedValue = DialogStateTracker.get(dialog.appId, GRAPPLE_FORM_KEY);
  if (storedValue !== undefined) {
    input.checked = storedValue;
  }
  input.addEventListener('change', () => {
    DialogStateTracker.set(dialog.appId, GRAPPLE_FORM_KEY, input.checked);
  });
  labelElement.textContent = ` ${labelText} `;
  labelElement.insertBefore(input, labelElement.firstChild);
  container.appendChild(labelElement);
  dialog.setPosition();
}

export async function addMetamagicCheckbox(dialog, html) {
  if (!(dialog instanceof pf1.applications.AttackDialog)) return;
  if (getActionItemType(dialog) !== "spell") return;

  const actor = getActionActor(dialog);
  const root = Array.isArray(html) ? html[0] : html?.[0] || html;
  if (!root) return;
  const form = root.querySelector?.('form') ?? root;
  if (!form) return;

  const metamagicSources = await getAvailableMetamagicSources(actor, { resolveEnglishNames: false });
  const filteredSources = filterMetamagicSourcesForDialog(dialog, metamagicSources);
  if (filteredSources.length) {
    renderMetamagicControls(dialog, form, filteredSources);
  }

  const isEnglish = (game?.i18n?.lang ?? "en").toLowerCase().startsWith("en");
  const canUseBabele = game?.modules?.get("babele")?.active;
  const shouldResolveEnglishName = !isEnglish && canUseBabele;
  if (shouldResolveEnglishName) {
    void refreshMetamagicControls(dialog, form, actor);
  }
}

function renderMetamagicControls(dialog, form, metamagicSources) {
  if (form.querySelector(`input[name="${METAMAGIC_FORM_KEY}"]`)) return;

  const container = getFlagsContainer(form);
  const labelElement = document.createElement('label');
  labelElement.classList.add('checkbox');
  const input = document.createElement('input');
  input.type = 'checkbox';
  input.name = METAMAGIC_FORM_KEY;
  const storedChecked = DialogStateTracker.get(dialog.appId, METAMAGIC_FORM_KEY);
  if (storedChecked !== undefined) {
    input.checked = storedChecked;
  }
  labelElement.textContent = ' Metamagic ';
  labelElement.insertBefore(input, labelElement.firstChild);
  container.appendChild(labelElement);

  const dataInput = document.createElement('input');
  dataInput.type = 'hidden';
  dataInput.name = METAMAGIC_NAMES_KEY;
  container.appendChild(dataInput);

  const optionsInput = document.createElement('input');
  optionsInput.type = 'hidden';
  optionsInput.name = METAMAGIC_OPTIONS_KEY;
  container.appendChild(optionsInput);

  const dropdown = document.createElement('details');
  dropdown.classList.add('metamagic-dropdown');
  dropdown.style.display = 'none';
  const summary = document.createElement('summary');
  summary.textContent = 'Select metamagic';
  dropdown.appendChild(summary);
  const listContainer = document.createElement('div');
  listContainer.classList.add('metamagic-options');
  listContainer.style.maxHeight = '160px';
  listContainer.style.overflowY = 'auto';
  listContainer.style.display = 'grid';
  listContainer.style.gridTemplateColumns = 'repeat(3, minmax(0, 1fr))';
  listContainer.style.columnGap = '8px';
  dropdown.appendChild(listContainer);
  container.appendChild(dropdown);

  updateMetamagicCheckboxOptions(dialog, listContainer, dataInput, optionsInput, metamagicSources);
  dropdown.style.display = input.checked ? '' : 'none';
  const storedOpen = DialogStateTracker.get(dialog.appId, METAMAGIC_DROPDOWN_KEY);
  if (storedOpen !== undefined) {
    dropdown.open = storedOpen;
  }

  input.addEventListener('change', () => {
    DialogStateTracker.set(dialog.appId, METAMAGIC_FORM_KEY, input.checked);
    dropdown.style.display = input.checked ? '' : 'none';
    updateMetamagicNames(dataInput, input.checked, listContainer);
    dialog.setPosition();
  });
  dropdown.addEventListener('toggle', () => {
    DialogStateTracker.set(dialog.appId, METAMAGIC_DROPDOWN_KEY, dropdown.open);
    dialog.setPosition();
  });

  dialog.setPosition();
}

function updateMetamagicNames(dataInput, isChecked, listContainer) {
  if (!isChecked) {
    dataInput.value = "";
    return;
  }

  const selections = Array.from(
    listContainer.querySelectorAll('input[type="checkbox"][data-meta-name]:checked')
  )
    .map((input) => input.dataset.metaName)
    .filter(Boolean);

  dataInput.value = selections.length ? JSON.stringify(selections) : "";
}

function updateMetamagicOptionsInput(optionsInput, options) {
  const hasOptions = options && Object.keys(options).length > 0;
  optionsInput.value = hasOptions ? JSON.stringify(options) : "";
}

function updateMetamagicCheckboxOptions(dialog, listContainer, dataInput, optionsInput, metamagicSources) {
  const storedSelections = DialogStateTracker.get(dialog.appId, METAMAGIC_SELECT_KEY) || [];
  const storedOptions = DialogStateTracker.get(dialog.appId, METAMAGIC_OPTIONS_KEY) || {};
  listContainer.replaceChildren();

  metamagicSources
    .slice()
    .sort((a, b) => a.label.localeCompare(b.label))
    .forEach((source) => {
      const labelElement = document.createElement('label');
      labelElement.classList.add('checkbox');

      const input = document.createElement('input');
      input.type = 'checkbox';
      input.value = source.item?.uuid ?? source.item?.id ?? source.label;
      input.dataset.type = source.type;
      input.dataset.metaName = source.metaName ?? source.label;
      input.checked = storedSelections.includes(input.value);

      input.addEventListener('change', () => {
        const checkedIds = Array.from(
          listContainer.querySelectorAll('input[type="checkbox"]:checked')
        ).map((checkbox) => checkbox.value);
        DialogStateTracker.set(dialog.appId, METAMAGIC_SELECT_KEY, checkedIds);
        updateMetamagicNames(dataInput, true, listContainer);
        updateMetamagicOptionsInput(optionsInput, DialogStateTracker.get(dialog.appId, METAMAGIC_OPTIONS_KEY));
      });

      if (input.dataset.metaName === "Reach Spell") {
        input.addEventListener('change', async () => {
          if (!input.checked) return;
          const rangeUnits = dialog?.action?.range?.units ?? "touch";
          const stepChoices = getReachSpellStepChoices(rangeUnits);
          if (!stepChoices.length) return;
          const steps = await promptReachSpellSteps(stepChoices);
          if (steps == null) {
            input.checked = false;
            input.dispatchEvent(new Event('change'));
            return;
          }
          const nextOptions = { ...(DialogStateTracker.get(dialog.appId, METAMAGIC_OPTIONS_KEY) || {}) };
          nextOptions.reachSpellSteps = steps;
          DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, nextOptions);
          updateMetamagicOptionsInput(optionsInput, nextOptions);
        });
      }

      if (input.dataset.metaName === "Heighten Spell") {
        input.addEventListener('change', async () => {
          if (!input.checked) {
            const nextOptions = { ...(DialogStateTracker.get(dialog.appId, METAMAGIC_OPTIONS_KEY) || {}) };
            delete nextOptions.heightenSpellLevel;
            DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, nextOptions);
            updateMetamagicOptionsInput(optionsInput, nextOptions);
            return;
          }
          const baseLevel = getSpellBaseLevel(dialog);
          const levelChoices = getHeightenSpellLevelChoices(baseLevel);
          if (!levelChoices.length) return;
          const level = await promptHeightenSpellLevel(levelChoices, baseLevel);
          if (level == null) {
            input.checked = false;
            input.dispatchEvent(new Event('change'));
            return;
          }
          const nextOptions = { ...(DialogStateTracker.get(dialog.appId, METAMAGIC_OPTIONS_KEY) || {}) };
          nextOptions.heightenSpellLevel = level;
          DialogStateTracker.set(dialog.appId, METAMAGIC_OPTIONS_KEY, nextOptions);
          updateMetamagicOptionsInput(optionsInput, nextOptions);
        });
      }

      labelElement.textContent = ` ${source.label} `;
      labelElement.insertBefore(input, labelElement.firstChild);
      listContainer.appendChild(labelElement);
    });

  updateMetamagicNames(dataInput, true, listContainer);
  updateMetamagicOptionsInput(optionsInput, storedOptions);
  dialog.setPosition();
}

function getReachSpellStepChoices(units) {
  const order = ["touch", "close", "medium", "long"];
  const normalized = units?.toString?.().toLowerCase() ?? "";
  const startIndex = order.indexOf(normalized);
  if (startIndex === -1) return [];
  const maxSteps = order.length - 1 - startIndex;
  return Array.from({ length: maxSteps }, (_, idx) => {
    const step = idx + 1;
    return {
      step,
      from: capitalizeRangeLabel(order[startIndex]),
      to: capitalizeRangeLabel(order[startIndex + step]),
    };
  });
}

function capitalizeRangeLabel(value) {
  if (!value) return "";
  return value.charAt(0).toUpperCase() + value.slice(1);
}

function promptReachSpellSteps(stepChoices) {
  return new Promise((resolve) => {
    const options = stepChoices
      .map((choice) => `<option value="${choice.step}">${choice.from} → ${choice.to}</option>`)
      .join("");
    const content = `
      <form>
        <div class="form-group">
          <label>Range steps</label>
          <select name="reachSteps">${options}</select>
        </div>
      </form>
    `;
    new Dialog({
      title: "Reach Spell",
      content,
      buttons: {
        ok: {
          label: "Apply",
          callback: (html) => {
            const value = Number(html.find('select[name="reachSteps"]').val());
            resolve(Number.isFinite(value) ? value : null);
          },
        },
        cancel: {
          label: "Cancel",
          callback: () => resolve(null),
        },
      },
      default: "ok",
      close: () => resolve(null),
    }).render(true);
  });
}

function getHeightenSpellLevelChoices(baseLevel) {
  const start = Number(baseLevel ?? 0);
  if (!Number.isFinite(start) || start >= 9) return [];
  return Array.from({ length: 9 - start }, (_, index) => {
    const level = start + index + 1;
    return { level, label: `Level ${level}` };
  });
}

function promptHeightenSpellLevel(levelChoices, baseLevel) {
  return new Promise((resolve) => {
    const options = levelChoices
      .map((choice) => `<option value="${choice.level}">${choice.label}</option>`)
      .join("");
    const content = `
      <form>
        <p style="margin: 0 0 8px 0;">Current level: <b>${baseLevel}</b></p>
        <div class="form-group">
          <label>Heighten to</label>
          <select name="heightenLevel">${options}</select>
        </div>
      </form>
    `;
    new Dialog({
      title: "Heighten Spell",
      content,
      buttons: {
        ok: {
          label: "Apply",
          callback: (html) => {
            const value = Number(html.find('select[name="heightenLevel"]').val());
            resolve(Number.isFinite(value) ? value : null);
          },
        },
        cancel: {
          label: "Cancel",
          callback: () => resolve(null),
        },
      },
      default: "ok",
      close: () => resolve(null),
    }).render(true);
  });
}

async function refreshMetamagicControls(dialog, form, actor) {
  const metamagicSources = await getAvailableMetamagicSources(actor, { resolveEnglishNames: true });
  const filteredSources = filterMetamagicSourcesForDialog(dialog, metamagicSources);
  if (!filteredSources.length) return;

  const listContainer = form.querySelector('.metamagic-options');
  const dataInput = form.querySelector(`input[name="${METAMAGIC_NAMES_KEY}"]`);
  const optionsInput = form.querySelector(`input[name="${METAMAGIC_OPTIONS_KEY}"]`);
  if (listContainer && dataInput && optionsInput) {
    updateMetamagicCheckboxOptions(dialog, listContainer, dataInput, optionsInput, filteredSources);
    dialog.setPosition();
    return;
  }

  renderMetamagicControls(dialog, form, filteredSources);
}

