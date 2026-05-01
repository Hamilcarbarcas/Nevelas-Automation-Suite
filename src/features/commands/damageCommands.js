
import { MODULE } from '../../common/module.js';
import {
  applyNasChatDamageButton,
  applyNasHeadlessDamage
} from '../automation/damage/systemApplyDamage.js';
import { toDamagePartModel } from '../automation/damage/instances.js';

export class DamageCommands {
  static i18n(path, data = null) {
    const key = `NAS.commands.damage.${path}`;
    return data ? game.i18n.format(key, data) : game.i18n.localize(key);
  }

  static initialize() {
    Hooks.on('chatMessage', this.handleChatCommand.bind(this));
    this.registerApplyButtonWrapper();
    
    window.ASDamage = {
      roll: this.macroRoll.bind(this)
    };
  }

  static registerApplyButtonWrapper() {
    if (!globalThis.libWrapper) return;
    libWrapper.register(
      MODULE.ID,
      "pf1.utils.chat.onButton",
      this.onChatButton.bind(this),
      libWrapper.MIXED
    );
  }

  static async macroRoll(options) {
    let formula = "";
    let damageTypes = [];
    let mixedDamage = false;
    
    if (typeof options === "string") {
      if (this.isMixedDamageFormat(options)) {
        return this.createMixedDamageRoll(options);
      }
      
      const typeMatch = options.match(/type\s*:\s*([^,]+)/i);
      const forMatch = options.match(/for\s*:\s*([^,]+)/i);
      
      if (typeMatch && forMatch) {
        damageTypes = typeMatch[1].split(/\s*\+\s*/).map(t => t.trim());
        formula = forMatch[1].trim();
      } else {
        ui.notifications.error(this.i18n("errors.invalidMacroFormat"));
        return null;
      }
    } else if (typeof options === "object") {
      if (options.components) {
        return this.createComponentDamageRoll(options.components);
      } else if (options.mixedDamage) {
        return this.createMixedDamageRoll(options.mixedDamage);
      }
      
      formula = options.formula || "";
      damageTypes = Array.isArray(options.damageTypes) ? options.damageTypes : [options.damageType || ""];
    } else {
      ui.notifications.error(this.i18n("errors.invalidMacroArgument"));
      return null;
    }
    
    if (!formula) {
      ui.notifications.error(this.i18n("errors.noFormulaProvided"));
      return null;
    }
    
    if (!damageTypes.length || !damageTypes[0]) {
      ui.notifications.warn(this.i18n("warnings.noDamageTypeProvidedUsingUntyped"));
      damageTypes = ["untyped"];
    }
    
    return this.createDamageRoll(formula, damageTypes);
  }
  
  static isMixedDamageFormat(str) {
    const mixedPattern = /(\d+d\d+|\d+)\s+([a-zA-Z]+)/;
    return mixedPattern.test(str);
  }
  
  static parseMixedDamageString(mixedStr) {
    const components = [];
    
    let inParentheses = 0;
    let currentComponent = "";
    
    for (let i = 0; i < mixedStr.length; i++) {
      const char = mixedStr[i];
      
      if (char === '(') inParentheses++;
      else if (char === ')') inParentheses--;
      
      if (char === '+' && inParentheses === 0) {
        components.push(currentComponent.trim());
        currentComponent = "";
      } else {
        currentComponent += char;
      }
    }
    
    if (currentComponent.trim()) {
      components.push(currentComponent.trim());
    }
    
    return components.map(component => {
      component = component.trim();
      if (component.startsWith('(') && component.endsWith(')')) {
        component = component.substring(1, component.length - 1).trim();
      }
      
      const formulaAndTypes = component.match(/^([\d\s\+\-\*\/\(\)d]+)\s+([a-zA-Z\s]+)$/);
      
      if (formulaAndTypes) {
        const formula = formulaAndTypes[1].trim();
        const damageTypesStr = formulaAndTypes[2].trim();
        
        const damageTypes = damageTypesStr.split(/\s+/).filter(type => type.length > 0);
        
        return {
          formula: formula,
          damageType: damageTypes[0] || "untyped" 
        };
      } else {
        return {
          formula: component.trim(),
          damageType: "untyped"
        };
      }
    });
  }
  
  static async createComponentDamageRoll(components) {
    try {
      if (components.length === 0) {
        ui.notifications.error(this.i18n("errors.noDamageComponentsProvided"));
        return null;
      }
      
      const normalizedComponents = components.map((component) => ({
        ...component,
        damageTypes: this.resolveDamageTypeIds(component.damageTypes)
      }));
      
      const damageTypeMap = this.getDamageTypes();
      
      const allDamageTypes = new Set();
      normalizedComponents.forEach(c => c.damageTypes.forEach(t => allDamageTypes.add(t)));
      const damageTypesArray = Array.from(allDamageTypes);
      
      const componentRolls = [];
      const damageInstances = [];
      let totalDamage = 0;
      
      for (const component of normalizedComponents) {
        const localizedComponentTypes = component.damageTypes.map(type => 
          damageTypeMap[type] || type 
        );
        
        const roll = new pf1.dice.DamageRoll(component.formula, {}, {
          damageType: component.damageTypes,
          type: "normal"
        });
        
        await roll.evaluate();
        
        for (const term of roll.terms) {
          if (term.class !== "OperatorTerm") {
            if (!term.options) term.options = {};
            term.options.flavor = localizedComponentTypes.join(", ");
            term.options.damageType = component.damageTypes;
          }
        }
        
        componentRolls.push(roll);
        damageInstances.push({
          types: component.damageTypes,
          value: roll.total,
          formula: String(roll.total)
        });
        totalDamage += roll.total;
      }
      
      const fullDamage = totalDamage;
      const halfDamage = Math.floor(totalDamage / 2);
      
      const buttonsHtml = `
        <div class="card-buttons flexrow">
          <button type="button" data-action="applyDamage" data-ratio="1" data-value="${fullDamage}" data-damage-types="${damageTypesArray.join(',')}">${this.i18n("chat.apply")}</button>
          <button type="button" data-action="applyDamage" data-ratio="0.5" data-value="${halfDamage}" data-damage-types="${damageTypesArray.join(',')}">${this.i18n("chat.applyHalf")}</button>
        </div>
      `;
      
      let content = `<div class="dice-roll">`;
      content += `<h2>${MODULE.SHORTNAME}</h2>`;
      content += `<div class="dice-result">`;
      
      content += `<div class="dice-tooltip">`;
      for (let i = 0; i < componentRolls.length; i++) {
        const roll = componentRolls[i];
        const rollContent = await roll.render();
        
        const $rollContent = $(rollContent);
        const $tooltip = $rollContent.find('.dice-tooltip');
        
        const $diceElements = $tooltip.find('.dice');
        if ($diceElements.length > 1) {
          const $firstDice = $diceElements.first();
          $diceElements.not(':first').each(function() {
            $(this).find('.dice-result').appendTo($firstDice);
            $(this).remove();
          });
        }
        
        const $partFormula = $tooltip.find('.part-formula');
        if ($partFormula.length > 0) {
          $partFormula.text(components[i].formula);
        }
        
        const $partTotals = $tooltip.find('.part-total');
        if ($partTotals.length > 0) {
          $partTotals.html(roll.total);
        }
        
        content += `\n  <section class=\"component-tooltip\">\n    <div class=\"component-label\"><strong>${this.i18n("dialogs.messages.componentLabel", { index: i + 1 })}</strong></div>\n    ${$tooltip.html()}\n  </section>\n`;
      }
      content += `</div>`; 
      
      content += `<h4 class="dice-total">${totalDamage}</h4>`;
      content += `</div>`; 
      content += `</div>`; 
      
      content += buttonsHtml;
      
      const messageData = {
        content: content,
        speaker: ChatMessage.getSpeaker(),
        rolls: componentRolls, 
        flags: {
          [MODULE.ID]: {
            damageRoll: true,
            damageTotal: fullDamage,
            damageTypes: damageTypesArray,
            damageInstances: damageInstances,
            componentRoll: true,
            source: "command",
            subject: {
              health: "damage"
            }
          }
        }
      };
      
      const message = await ChatMessage.create(messageData);

      return message;
    } catch (error) {
      console.error(`${MODULE.ID} | Error creating component damage roll`, error);
      ui.notifications.error(this.i18n("errors.errorCreatingDamageRoll"));
      return null;
    }
  }

  static async createMixedDamageRoll(mixedDamageString) {
    try {
      const damageComponents = this.parseMixedDamageString(mixedDamageString);
      
      if (damageComponents.length === 0) {
        ui.notifications.error(this.i18n("errors.failedParseMixedDamageFormula"));
        return null;
      }
      
      const components = damageComponents.map(component => ({
        formula: component.formula,
        damageTypes: [component.damageType]
      }));
      
      return this.createComponentDamageRoll(components);
    } catch (error) {
      console.error(`${MODULE.ID} | Error creating mixed damage roll`, error);
      ui.notifications.error(this.i18n("errors.errorCreatingMixedDamageRoll"));
      return null;
    }
  }

  static getDamageTypes() {
    const damageTypes = pf1.registry.damageTypes;
    const alignments = pf1.config.damageResistances;
    
    const allTypes = {};
    
    if (damageTypes && damageTypes instanceof Map) {
      for (const [key, value] of damageTypes.entries()) {
        if (value.name) {
          allTypes[key] = value.name;
        }
      }
    }
    
    if (alignments) {
      for (const [key, value] of Object.entries(alignments)) {
        if (!allTypes[key]) {
          allTypes[key] = value;
        }
      }
    }
    
    return allTypes;
  }

  static resolveDamageTypeIds(damageTypes) {
    const raw = Array.isArray(damageTypes) ? damageTypes : [damageTypes];
    const resolved = [];
    for (const typeRef of raw) {
      if (!typeRef || typeof typeRef !== "string") continue;
      const id = this.normalizeDamageTypeId(typeRef);
      if (id) resolved.push(id);
    }
    return resolved.length ? Array.from(new Set(resolved)) : ["untyped"];
  }

  static normalizeDamageTypeId(typeRef) {
    const needle = String(typeRef).trim().toLowerCase();
    if (!needle) return null;

    const reg = pf1?.registry?.damageTypes;
    if (reg?.get) {
      const direct = reg.get(typeRef);
      if (direct) return typeRef;
      for (const [key, value] of reg.entries()) {
        const name = value?.name?.toLowerCase();
        const shortName = value?.shortName?.toLowerCase();
        if (needle === String(key).toLowerCase() || needle === name || needle === shortName) {
          return key;
        }
      }
    }

    const alignments = pf1?.config?.damageResistances || {};
    for (const [key, label] of Object.entries(alignments)) {
      if (needle === String(key).toLowerCase()) return key;
      if (needle === String(label).toLowerCase()) return key;
    }

    return null;
  }

  static async onChatButton(wrapped, message, elementObject) {
    const nasChatDamage = await applyNasChatDamageButton(message, elementObject);
    if (nasChatDamage?.handled) return false;

    const button = elementObject?.currentTarget ?? elementObject?.target;
    const actionName = button?.dataset?.action;
    if (actionName !== "applyDamage") return wrapped(message, elementObject);

    const flags = message?.flags?.[MODULE.ID];
    if (!flags?.damageRoll || flags?.source !== "command") {
      return wrapped(message, elementObject);
    }

    if (!button) return wrapped(message, elementObject);

    const rawInstances = Array.isArray(flags.damageInstances) ? flags.damageInstances : [];
    if (!rawInstances.length) return wrapped(message, elementObject);

    const ratio = Number(button.dataset.ratio);
    const appliedRatio = (Number.isFinite(ratio) && ratio > 0) ? ratio : 1;

    const instances = (foundry.utils?.deepClone ? foundry.utils.deepClone(rawInstances) : rawInstances)
      .map((inst) => {
        const rawTypes = inst?.types ?? inst?.typeIds ?? [];
        const typeList = rawTypes instanceof Set
          ? Array.from(rawTypes)
          : (Array.isArray(rawTypes) ? rawTypes : [rawTypes]);
        const resolvedTypeIds = this.resolveDamageTypeIds(typeList.filter((entry) => typeof entry === "string"));
        const numericValue = Number(inst?.value ?? inst?.total ?? inst?.formula) || 0;

        return toDamagePartModel({
          types: resolvedTypeIds,
          value: numericValue,
          formula: String(numericValue)
        });
      });

    const totalValue = instances.reduce((sum, inst) => {
      return sum + (Number(inst?.value ?? inst?.total ?? inst?.formula) || 0);
    }, 0);

    if (!Number.isFinite(totalValue) || totalValue === 0) return false;

    const ev = elementObject?.event ?? elementObject?.originalEvent ?? null;
    const applyOptions = {
      instances,
      ratio: appliedRatio,
      message,
      element: button,
      event: ev
    };

    const applied = await applyNasHeadlessDamage(totalValue, applyOptions);
    if (!applied?.handled) {
      pf1?.documents?.actor?.ActorPF?.applyDamage?.(totalValue, applied?.options ?? applyOptions);
    }
    return false;
  }

  static showDamageDialog() {
    const damageTypes = this.getDamageTypes();
    
    const sortedTypes = Object.entries(damageTypes)
      .sort((a, b) => a[1].localeCompare(b[1]))
      .reduce((obj, [key, value]) => {
        obj[key] = value;
        return obj;
      }, {});
    
    let content = `
      <form>
        <div class="form-group">
          <label>${this.i18n("dialogs.labels.damageComponents")}</label>
          <div id="damage-components-list">
            <!-- Damage components will be added here -->
            <div class="empty-message">${this.i18n("dialogs.messages.emptyComponents")}</div>
          </div>
        </div>
        
        <div class="form-group">
          <button type="button" id="add-damage-component" class="damage-component-add-btn">
            <i class="fas fa-plus"></i> ${this.i18n("dialogs.buttons.addRoll")}
          </button>
        </div>
        
        <div class="components-info" style="font-size: 0.9em; color: #777; margin-top: 5px;">
          ${this.i18n("dialogs.messages.componentsInfo")}
        </div>
      </form>
    `;
    
    const damageComponents = [];
    
    const mainDialog = new Dialog({
      title: this.i18n("dialogs.mainTitle", { shortName: MODULE.SHORTNAME }),
      content: content,
      buttons: {
        roll: {
          icon: '<i class="fas fa-dice-d20"></i>',
          label: this.i18n("dialogs.buttons.rollDamage"),
          callback: (html) => {
            if (damageComponents.length === 0) {
              ui.notifications.error(this.i18n("errors.addAtLeastOneDamageComponent"));
              return;
            }
            
            this.createComponentDamageRoll(damageComponents);
          }
        },
        cancel: {
          icon: '<i class="fas fa-times"></i>',
          label: this.i18n("dialogs.buttons.cancel")
        }
      },
      default: "roll",
      render: (html) => {
        const style = document.createElement('style');
        style.textContent = `
          .damage-components-list {
            display: flex;
            flex-direction: column;
            gap: 10px;
            margin-bottom: 10px;
          }
          .damage-component {
            border: 1px solid #ccc;
            border-radius: 4px;
            padding: 10px;
            margin-bottom: 10px;
            background: rgba(0, 0, 0, 0.05);
            position: relative;
          }
          .damage-component-formula {
            font-family: monospace;
            font-weight: bold;
            margin-bottom: 5px;
          }
          .damage-component-type {
            font-style: italic;
            color: #555;
          }
          .damage-component-remove {
            position: absolute;
            top: 5px;
            right: 5px;
            cursor: pointer;
            color: #a00;
            font-size: 12px;
          }
          .damage-component-add-btn {
            width: 100%;
            margin-top: 5px;
            background: #f0f0f0;
            border: 1px solid #ccc;
            padding: 5px;
            border-radius: 4px;
            cursor: pointer;
          }
          .damage-component-add-btn:hover {
            background: #e8e8e8;
          }
          .empty-message {
            font-style: italic;
            color: #777;
            padding: 10px;
            text-align: center;
          }
          .checkbox-grid {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 5px;
            max-height: 200px;
            overflow-y: auto;
            margin-top: 10px;
            padding: 5px;
            border: 1px solid #ddd;
            border-radius: 3px;
          }
          .checkbox-item {
            display: flex;
            align-items: center;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
          }
          .checkbox-item label {
            margin-left: 5px;
            overflow: hidden;
            text-overflow: ellipsis;
            font-weight: normal;
          }
        `;
        html.find('form').prepend(style);
        
        html.find('#add-damage-component').click(() => {
          this.showDamageComponentDialog(sortedTypes, (component) => {
            damageComponents.push(component);
            
            updateComponentsList();
          });
        });
        
        const updateComponentsList = () => {
          const list = html.find('#damage-components-list');
          
          list.empty();
          
          if (damageComponents.length === 0) {
            list.append(`<div class="empty-message">${this.i18n("dialogs.messages.emptyComponents")}</div>`);
            const dialog = html.closest('.app');
            dialog.css({
              'height': '210px',
              'min-height': '',
              'max-height': '',
              'overflow-y': '',
              'width': '500px',
              'min-width': '400px',
              'max-width': '500px'
            });
            return;
          } else {
            const dialog = html.closest('.app');
            dialog.css({
              'height': '',
              'min-height': '210px',
              'max-height': '420px',
              'overflow-y': 'auto',
              'width': '500px',
              'min-width': '400px',
              'max-width': '500px'
            });
          }
          
          damageComponents.forEach((component, index) => {
            const typeNames = component.damageTypes.map(type => sortedTypes[type] || type);
            const typeLabel = typeNames.length === 1
              ? this.i18n("dialogs.messages.componentTypeSingle")
              : this.i18n("dialogs.messages.componentTypePlural");
            const componentHtml = `
              <div class="damage-component" data-index="${index}">
                <div class="damage-component-formula">${component.formula}</div>
                <div class="damage-component-type">${typeLabel} ${typeNames.join(", ")}</div>
                <a class="damage-component-remove"><i class="fas fa-times"></i></a>
              </div>
            `;
            list.append(componentHtml);
          });
          
          list.find('.damage-component-remove').click(function() {
            const index = $(this).parent().data('index');
            damageComponents.splice(index, 1);
            updateComponentsList();
          });
        };
        
        updateComponentsList();
      },
      width: 400  
    }).render(true);
  }
  
  static showDamageComponentDialog(sortedTypes, callback) {
    let content = `
      <form>
        <div class="form-group">
          <label for="component-formula">${this.i18n("dialogs.labels.damageFormula")}</label>
          <input type="text" id="component-formula" name="formula" placeholder="${this.i18n("dialogs.placeholders.formula")}" style="width: 100%;">
        </div>
        
        <div class="form-group">
          <label>${this.i18n("dialogs.labels.damageTypes")}</label>
        </div>
        
        <div class="checkbox-grid">
    `;
    
    for (const [key, name] of Object.entries(sortedTypes)) {
      content += `
        <div class="checkbox-item">
          <input type="checkbox" id="component-type-${key}" name="damageTypes" value="${key}">
          <div class="checkbox-label-container">
            <label for="component-type-${key}" title="${name}">${name}</label>
          </div>
        </div>
      `;
    }
    
    content += `
        </div>
      </form>
    `;
    
    const dialog = new Dialog({
      title: this.i18n("dialogs.addDamageRollTitle"),
      content: content,
      buttons: {
        add: {
          icon: '<i class="fas fa-plus"></i>',
          label: this.i18n("dialogs.buttons.add"),
          callback: (html) => {
            const formula = html.find('#component-formula').val().trim();
            if (!formula) {
              ui.notifications.error(this.i18n("errors.enterDamageFormula"));
              return;
            }
            
            const selectedTypes = [];
            html.find('input[name="damageTypes"]:checked').each((i, cb) => {
              selectedTypes.push(cb.value);
            });
            
            if (selectedTypes.length === 0) {
              ui.notifications.error(this.i18n("errors.selectAtLeastOneDamageType"));
              return;
            }
            
            callback({
              formula: formula,
              damageTypes: selectedTypes
            });
          }
        },
        cancel: {
          icon: '<i class="fas fa-times"></i>',
          label: this.i18n("dialogs.buttons.cancel")
        }
      },
      default: "add",
      render: (html) => {
        const style = document.createElement('style');
        style.textContent = `
          .checkbox-item {
            display: flex;
            flex-direction: column;
            align-items: center;
            text-align: center;
            margin-bottom: 10px;
          }
          .checkbox-label-container {
            margin-top: 3px;
            text-align: center;
          }
          .checkbox-grid {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 10px;
            max-height: 250px;
            overflow-y: auto;
            padding: 10px;
            border: 1px solid #ddd;
            border-radius: 3px;
          }
        `;
        html.find('form').append(style);
        
        const dialog = html.closest('.app');
        dialog.css('min-width', '400px');
        dialog.css('max-width', '500px');
      },
      width: 400  
    }).render(true);
  }

  static processDamageCommand(args, chatData) {
    const argsParts = args.trim().split(/\s+/);
    if (argsParts.length > 0) {
      const firstArg = argsParts[0].toLowerCase();
      
      if (firstArg === 'macro') {
        this.showMacroCreationDialog();
        return false;
      }
    }
    
    if (!args || args.trim().length === 0) {
      this.showDamageDialog();
      return false;
    }
    
    if (this.isMixedDamageFormat(args)) {
      this.createMixedDamageRoll(args);
      return false;
    }
    
    const parts = args.trim().split(/\s+/);
    
    if (parts.length < 2) {
      this.showFormatError();
      return false;
    }
    
    const formula = parts[0];
    
    const damageTypes = parts.slice(1);

    this.createDamageRoll(formula, damageTypes);
    return false;
  }

  static async createDamageRoll(formula, damageTypes) {
    try {
      const resolvedTypes = this.resolveDamageTypeIds(damageTypes);
      
      const roll = new pf1.dice.DamageRoll(formula, {}, {
        damageType: resolvedTypes,
        type: "normal"
      });
      
      await roll.evaluate();
      
      const damageTypeString = resolvedTypes.join(", ");
      for (const term of roll.terms) {
        if (term.class !== "OperatorTerm") {
          if (!term.options) term.options = {};
          term.options.flavor = damageTypeString;
        }
      }
      
      const fullDamage = roll.total;
      const halfDamage = Math.floor(roll.total / 2);
      
      const buttonsHtml = `
        <div class="card-buttons flexrow">
          <button type="button" data-action="applyDamage" data-ratio="1" data-value="${fullDamage}" data-damage-types="${resolvedTypes.join(',')}">${this.i18n("chat.apply")}</button>
          <button type="button" data-action="applyDamage" data-ratio="0.5" data-value="${halfDamage}" data-damage-types="${resolvedTypes.join(',')}">${this.i18n("chat.applyHalf")}</button>
        </div>
      `;
      
      const rollContent = await roll.render();
      
      const messageData = {
        flavor: this.i18n("chat.flavorDamage", { types: resolvedTypes.join(", ") }),
        content: rollContent + buttonsHtml,
        speaker: ChatMessage.getSpeaker(),
        rolls: [roll],
        flags: {
          [MODULE.ID]: {
            damageRoll: true,
            damageTotal: roll.total,
            damageTypes: resolvedTypes,
            damageInstances: [{
              types: resolvedTypes,
              value: roll.total,
              formula: String(roll.total)
            }],
            source: "command",
            subject: {
              health: "damage"
            }
          }
        }
      };
      
      const message = await ChatMessage.create(messageData);

      return message;
    } catch (error) {
      console.error(`${MODULE.ID} | Error creating damage roll`, error);
      this.showFormatError();
    }
  }

  static showFormatError() {
    ui.notifications.error(this.i18n("errors.invalidAdFormat"));
    ui.notifications.info(this.i18n("info.adExamples"));
  }

  static showMacroCreationDialog() {
    const damageTypes = this.getDamageTypes();
    
    const sortedTypes = Object.entries(damageTypes)
      .sort((a, b) => a[1].localeCompare(b[1]))
      .reduce((obj, [key, value]) => {
        obj[key] = value;
        return obj;
      }, {});
    
    let content = `
      <form>
        <div class="form-group">
          <label for="macro-name">${this.i18n("dialogs.labels.macroName")}</label>
          <input type="text" id="macro-name" name="macro-name" placeholder="${this.i18n("dialogs.placeholders.macroName", { shortName: MODULE.SHORTNAME })}" style="width: 100%;">
        </div>
        
        <div class="form-group">
          <label>${this.i18n("dialogs.labels.damageComponents")}</label>
          <div id="damage-components-list">
            <!-- Damage components will be added here -->
            <div class="empty-message">${this.i18n("dialogs.messages.emptyComponents")}</div>
          </div>
        </div>
        
        <div class="form-group">
          <button type="button" id="add-damage-component" class="damage-component-add-btn">
            <i class="fas fa-plus"></i> ${this.i18n("dialogs.buttons.addRoll")}
          </button>
        </div>
        
        <div class="components-info" style="font-size: 0.9em; color: #777; margin-top: 5px;">
          ${this.i18n("dialogs.messages.componentsInfo")}
        </div>
      </form>
    `;
    
    const damageComponents = [];
    
    const mainDialog = new Dialog({
      title: this.i18n("dialogs.createMacroTitle"),
      content: content,
      buttons: {
        create: {
          icon: '<i class="fas fa-save"></i>',
          label: this.i18n("dialogs.buttons.createMacro"),
          callback: (html) => {
            if (damageComponents.length === 0) {
              ui.notifications.error(this.i18n("errors.addAtLeastOneDamageComponent"));
              return;
            }
            
            let macroName = html.find('#macro-name').val().trim();
            if (!macroName) {
              macroName = this.getAvailableMacroName();
            }
            
            this.createDamageMacro(macroName, damageComponents);
          }
        },
        cancel: {
          icon: '<i class="fas fa-times"></i>',
          label: this.i18n("dialogs.buttons.cancel")
        }
      },
      default: "create",
      render: (html) => {
        const style = document.createElement('style');
        style.textContent = `
          .damage-components-list {
            display: flex;
            flex-direction: column;
            gap: 10px;
            margin-bottom: 10px;
          }
          .damage-component {
            border: 1px solid #ccc;
            border-radius: 4px;
            padding: 10px;
            margin-bottom: 10px;
            background: rgba(0, 0, 0, 0.05);
            position: relative;
          }
          .damage-component-formula {
            font-family: monospace;
            font-weight: bold;
            margin-bottom: 5px;
          }
          .damage-component-type {
            font-style: italic;
            color: #555;
          }
          .damage-component-remove {
            position: absolute;
            top: 5px;
            right: 5px;
            cursor: pointer;
            color: #a00;
            font-size: 12px;
          }
          .damage-component-add-btn {
            width: 100%;
            margin-top: 5px;
            background: #f0f0f0;
            border: 1px solid #ccc;
            padding: 5px;
            border-radius: 4px;
            cursor: pointer;
          }
          .damage-component-add-btn:hover {
            background: #e8e8e8;
          }
          .empty-message {
            font-style: italic;
            color: #777;
            padding: 10px;
            text-align: center;
          }
          .checkbox-grid {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 5px;
            max-height: 200px;
            overflow-y: auto;
            margin-top: 10px;
            padding: 5px;
            border: 1px solid #ddd;
            border-radius: 3px;
          }
          .checkbox-item {
            display: flex;
            align-items: center;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
          }
          .checkbox-item label {
            margin-left: 5px;
            overflow: hidden;
            text-overflow: ellipsis;
            font-weight: normal;
          }
        `;
        html.find('form').prepend(style);
        
        html.find('#add-damage-component').click(() => {
          this.showDamageComponentDialog(sortedTypes, (component) => {
            damageComponents.push(component);
            
            updateComponentsList();
          });
        });
        
        const updateComponentsList = () => {
          const list = html.find('#damage-components-list');
          
          list.empty();
          
          if (damageComponents.length === 0) {
            list.append(`<div class="empty-message">${this.i18n("dialogs.messages.emptyComponents")}</div>`);
            const dialog = html.closest('.app');
            dialog.css({
              'height': '235px',
              'min-height': '',
              'max-height': '',
              'overflow-y': '',
              'width': '500px',
              'min-width': '400px',
              'max-width': '500px'
            });
            return;
          } else {
            const dialog = html.closest('.app');
            dialog.css({
              'height': '',
              'min-height': '235px',
              'max-height': '470px',
              'overflow-y': 'auto',
              'width': '500px',
              'min-width': '400px',
              'max-width': '500px'
            });
          }
          
          damageComponents.forEach((component, index) => {
            const typeNames = component.damageTypes.map(type => sortedTypes[type] || type);
            const typeLabel = typeNames.length === 1
              ? this.i18n("dialogs.messages.componentTypeSingle")
              : this.i18n("dialogs.messages.componentTypePlural");
            const componentHtml = `
              <div class="damage-component" data-index="${index}">
                <div class="damage-component-formula">${component.formula}</div>
                <div class="damage-component-type">${typeLabel} ${typeNames.join(", ")}</div>
                <a class="damage-component-remove"><i class="fas fa-times"></i></a>
              </div>
            `;
            list.append(componentHtml);
          });
          
          list.find('.damage-component-remove').click(function() {
            const index = $(this).parent().data('index');
            damageComponents.splice(index, 1);
            updateComponentsList();
          });
        };
        
        updateComponentsList();
      },
      width: 400  
    }).render(true);
  }
  
  static getAvailableMacroName() {
      const baseName = this.i18n("dialogs.placeholders.macroName", { shortName: MODULE.SHORTNAME });
    let counter = 1;
    let name = baseName;
    
    while (game.macros.find(m => m.name === name)) {
      name = `${baseName} ${counter}`;
      counter++;
    }
    
    return name;
  }
  
  static createDamageMacro(name, components) {
    try {
      const macroCommand = `
ASDamage.roll({
  components: ${JSON.stringify(components, null, 2)}
});`;
      
      Macro.create({
        name: name,
        type: "script",
        img: "icons/svg/fire.svg", 
        command: macroCommand,
        flags: {
          [MODULE.ID]: {
            isDamageMacro: true
          }
        }
      }).then(macro => {
        ui.notifications.info(this.i18n("info.macroCreated", { name }));
      });
    } catch (error) {
      console.error(`${MODULE.ID} | Error creating macro`, error);
      ui.notifications.error(this.i18n("errors.errorCreatingDamageMacro"));
    }
  }

  static handleChatCommand(log, message, chatData) {
    if (!message.startsWith('/')) return;
    
    const command = message.split(' ')[0].substring(1);
    const args = message.substring(command.length + 2);
    
    if (command === 'as') {
      try {
        return this.processDamageCommand(args, chatData);
      } catch (error) {
        console.error(`${MODULE.ID} | Error processing command`, error);
      }
      return false;
    }
  }
} 