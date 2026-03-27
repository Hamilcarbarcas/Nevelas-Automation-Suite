import { MODULE } from "./module.js";
import { DamagePriorityForm } from "./settings/damagePriorityForm.js";
import { DamageTypeFormApplication } from "./settings/damageTypeForm.js";
import { TranslationForm } from "./settings/translationForm.js";
import { BuffCompendiaSelector } from "./settings/buffCompendiaSelector.js";
import { ModifierNameSettingsForm } from "./settings/modifierNameSettingsForm.js";
import { VariantMappingManager } from "./settings/variantMappingManager.js";
import { MigrationSettingsMenu } from "./migrations.js";

export const damageConfig = {
    weaponDamageTypes: [],
    additionalPhysicalDamageTypes: []
};

function registerSettings() {
    Handlebars.registerHelper('colorStyle', function(color) {
        return new Handlebars.SafeString(`style="color: ${color};"`);
    });

    Handlebars.registerHelper('includes', function(array, value) {
        if (!array) return false;
        return array.includes(value);
    });

    const stylesheets = [
        'src/templates/css/damage-settings-form.css',
        'src/templates/css/nas-module-settings.css'
    ];
    stylesheets.forEach(stylesheet => {
        const link = document.createElement('link');
        link.rel = 'stylesheet';
        link.type = 'text/css';
        link.href = `modules/${MODULE.ID}/${stylesheet}`;
        document.head.appendChild(link);
    });

    game.settings.register(MODULE.ID, "massiveDamage", {
        name: game.i18n.localize("NAS.settings.massiveDamage.name"),
        hint: game.i18n.localize("NAS.settings.massiveDamage.hint"),
        default: false,
        scope: "world",
        type: Boolean,
        config: true
    });

    game.settings.register(MODULE.ID, "metamagicCastTimeRule", {
        name: game.i18n.localize("NAS.settings.metamagicCastTimeRule.name"),
        hint: game.i18n.localize("NAS.settings.metamagicCastTimeRule.hint"),
        scope: "world",
        config: true,
        type: String,
        choices: {
            always: game.i18n.localize("NAS.settings.metamagicCastTimeRule.choices.always"),
            standard: game.i18n.localize("NAS.settings.metamagicCastTimeRule.choices.standard")
        },
        default: "always"
    });

    game.settings.register(MODULE.ID, "persistentSpellTargetMode", {
        name: game.i18n.localize("NAS.settings.persistentSpellTargetMode.name"),
        hint: game.i18n.localize("NAS.settings.persistentSpellTargetMode.hint"),
        scope: "world",
        config: true,
        type: String,
        choices: {
            current: game.i18n.localize("NAS.settings.persistentSpellTargetMode.choices.current"),
            selected: game.i18n.localize("NAS.settings.persistentSpellTargetMode.choices.selected"),
            message: game.i18n.localize("NAS.settings.persistentSpellTargetMode.choices.message")
        },
        default: "current"
    });

    game.settings.register(MODULE.ID, "saveRollTokenInteraction", {
        name: game.i18n.localize("NAS.settings.saveRollTokenInteraction.name"),
        hint: game.i18n.localize("NAS.settings.saveRollTokenInteraction.hint"),
        scope: "world",
        config: true,
        type: Boolean,
        default: true
    });

    game.settings.register(MODULE.ID, "damageTypePriority", {
        name: game.i18n.localize("NAS.settings.damageTypePriority.name"),
        hint: game.i18n.localize("NAS.settings.damageTypePriority.hint"),
        default: JSON.stringify([[], ["magic"], [], ["alchemicalSilver", "coldIron", "mithral", "nexavarianSteel", "sunsilver"], ["adamantine"], ["lawful", "chaotic", "good", "evil"], ["epic"]]),
        scope: "world",
        type: String,
        config: false
    });

    game.settings.registerMenu(MODULE.ID, "damageTypePriorityMenu", {
        name: game.i18n.localize("NAS.settings.damageTypePriorityMenu.name"),
        label: game.i18n.localize("NAS.settings.damageTypePriorityMenu.label"),
        hint: game.i18n.localize("NAS.settings.damageTypePriorityMenu.hint"),
        icon: "fas fa-cogs",
        type: DamagePriorityForm,
        restricted: true
    });

    game.settings.registerMenu(MODULE.ID, "customSetting", {
        name: game.i18n.localize("NAS.damageTypes.title"),
        label: game.i18n.localize("NAS.settings.customSetting.label"),
        hint: game.i18n.localize("NAS.settings.customSetting.hint"),
        icon: "fas fa-cogs",
        type: DamageTypeFormApplication,
        restricted: true
    });

    game.settings.register(MODULE.ID, 'customDamageTypes', {
        scope: 'world',
        config: false,
        type: Array,
        default: [],
        requiresReload: true 
    });

    game.settings.register(MODULE.ID, "translations", {
        name: "Translation Settings",
        scope: "world",
        config: false,
        type: Object,
        default: {
            hardness: "",
            construct: "",
            undead: ""
        }
    });

    game.settings.register(MODULE.ID, "metamagicNameCache", {
        scope: "world",
        config: false,
        type: Object,
        default: {}
    });

    game.settings.registerMenu(MODULE.ID, "translationMenu", {
        name: "Configure Translations",
        label: "Translations",
        hint: "Set the translations for various terms for the module to recognize.",
        icon: "fas fa-language",
        type: TranslationForm,
        restricted: true
    });
    
    if (!game.settings.settings.has(`${MODULE.ID}.migrationVersion`)) {
        game.settings.register(MODULE.ID, "migrationVersion", {
            name: "Migration Version",
            scope: "world",
            config: false,
            type: String,
            default: ""
        });
    }

    if (!game.settings.settings.has(`${MODULE.ID}.migrationTool`)) {
        game.settings.registerMenu(MODULE.ID, "migrationTool", {
            name: "Run Legacy Migration",
            label: "Open Migration Tool",
            hint: "Copy legacy module data into this module's settings and flags.",
            icon: "fas fa-database",
            type: MigrationSettingsMenu,
            restricted: true
        });
    }
}

function syncWeaponDamageTypes() {
    const damageTypePriority = JSON.parse(game.settings.get(MODULE.ID, "damageTypePriority"));
    const materialTypes = pf1.registry.materials;
    const alignments = pf1.config.damageResistances;

    const weaponDamageTypes = damageTypePriority.map(priorityLevel => {
        return priorityLevel.map(type => {
            if (!type || type.trim() === '') {
                return null;
            }
            const material = materialTypes.find(m => m.name === type);
            const alignmentKey = Object.keys(alignments).find(key => alignments[key] === type);

            if (material) {
                return material.id;
            } else if (alignmentKey) {
                return alignmentKey;
            } else {
                return type.toLowerCase();
            }
        }).filter(type => type !== null);
    });
    damageConfig.weaponDamageTypes = weaponDamageTypes;
    pf1.registry.damageTypes.forEach(damageType => {
        if (['slashing', 'bludgeoning', 'piercing'].includes(damageType.id)) {
            damageConfig.additionalPhysicalDamageTypes.push(damageType.id);
        }
    });
}

export function populateDefaultTypes() {
    return new Promise(async (resolve, reject) => {
        try {
            const materialTypes = pf1.registry.materialTypes;
            const alignments = Object.keys(pf1.config.damageResistances);

            const priorityLevels = {
                1: [],
                2: [],
                3: [],
                4: [],
                5: alignments.map(key => pf1.config.damageResistances[key]), 
                6: []
            };

            materialTypes.forEach(material => {
                let targetArray;

                switch (material.id) {
                    case 'magic':
                        targetArray = priorityLevels[1];
                        break;
                    case 'coldIron':
                    case 'alchemicalSilver':
                        targetArray = priorityLevels[3];
                        break;
                    case 'adamantine':
                        targetArray = priorityLevels[4];
                        break;
                    case 'epic':
                        targetArray = priorityLevels[6];
                        break;
                }

                if (material.treatedAs) {
                    switch (material.treatedAs) {
                        case 'magic':
                            targetArray = priorityLevels[1];
                            break;
                        case 'coldIron':
                        case 'alchemicalSilver':
                            targetArray = priorityLevels[3];
                            break;
                        case 'adamantine':
                            targetArray = priorityLevels[4];
                            break;
                        case 'epic':
                            targetArray = priorityLevels[6];
                            break;
                    }
                }

                if (targetArray && !targetArray.includes(material.name)) { 
                    targetArray.push(material.name);
                }
            });

            damageConfig.weaponDamageTypes = [
                [], 
                ...Object.values(priorityLevels)
            ];

            await game.settings.set(MODULE.ID, "damageTypePriority", JSON.stringify(damageConfig.weaponDamageTypes));
            resolve();
        } catch (error) {
            console.error("Error populating default types:", error);
            reject(error);
        }
    });
}

async function handleReadyHook() {
    console.log(damageConfig);

    const migrationKey = `migrationVersion`;
    const currentVersion = game.modules.get(MODULE.ID).version;
    let previousMigrationVersion;

    try {
        previousMigrationVersion = game.settings.get(MODULE.ID, migrationKey);
    } catch (e) {
        previousMigrationVersion = "0.0.0";
    }

    if (compareVersions(currentVersion, previousMigrationVersion) > 0) {
        await performMigration(); 
        await game.settings.set(MODULE.ID, migrationKey, currentVersion); 
    }
    
    const customDamageTypes = game.settings.get(MODULE.ID, "customDamageTypes");
    
    customDamageTypes.forEach(damageType => {
        const { value } = damageType;

        if (!["physical", "energy", "misc"].includes(value.category.toLowerCase())) {
            const localizationKey = `PF1.DamageTypeCategory.${value.category.toLowerCase()}`;
            if (!game.i18n.translations.PF1.DamageTypeCategory) {
                game.i18n.translations.PF1.DamageTypeCategory = {};
            }

            const capitalizedCategory = value.category
                .split(/[\s-]/) 
                .map(word => word.charAt(0).toUpperCase() + word.slice(1))
                .join(' ')
                .replace(/\b([A-Za-z]+)-([A-Za-z]+)\b/g, (match, p1, p2) => `${p1}-${p2.charAt(0).toUpperCase()}${p2.slice(1)}`);

            game.i18n.translations.PF1.DamageTypeCategory[value.category.toLowerCase()] = capitalizedCategory;
        }
    });
}

function handleRegistryHook(registry) {
    registerSettings();
    const customDamageTypes = game.settings.get(MODULE.ID, "customDamageTypes");

    customDamageTypes.forEach(damageType => {
        const { key, value } = damageType;

        if (!["physical", "energy", "misc"].includes(value.category.toLowerCase().trim())) {
            registry.constructor.CATEGORIES.push(value.category);
        }

        registry.register(MODULE.ID, key, value);
    });
}

function handleSetupHook() {
    syncWeaponDamageTypes();
}

function compareVersions(v1, v2) {
    const [major1, minor1 = 0, patch1 = 0] = v1.split('.').map(Number);
    const [major2, minor2 = 0, patch2 = 0] = v2.split('.').map(Number);

    if (major1 > major2) return 1;
    if (major1 < major2) return -1;
    if (minor1 > minor2) return 1;
    if (minor1 < minor2) return -1;
    if (patch1 > patch2) return 1;
    if (patch1 < patch2) return -1;

    return 0; 
}

async function performMigration() {
    const customDamageTypes = game.settings.get(MODULE.ID, "customDamageTypes");

    if (Array.isArray(customDamageTypes) && customDamageTypes.length > 0) {
        const damageTypesToReRegister = [];

        customDamageTypes.forEach(damageType => {
            const flags = damageType.value.flags[MODULE.ID];

            if (typeof flags.abilities === "object") {
                damageTypesToReRegister.push(damageType);

                const abilityKeys = Object.keys(flags.abilities);
                flags.abilities = abilityKeys.join(",");
            }
        });

        if (damageTypesToReRegister.length > 0) {
            ui.notifications.info("Starting migration for custom damage types...");

            unregisterDamageTypes(damageTypesToReRegister);

            await game.settings.set(MODULE.ID, "customDamageTypes", customDamageTypes);

            reRegisterDamageTypes(damageTypesToReRegister);

            ui.notifications.info("Migration completed successfully!");
        }
    }
}

function unregisterDamageTypes(damageTypesToUnregister) {
    const registry = pf1.registry.damageTypes;

    damageTypesToUnregister.forEach(damageType => {
        const { key } = damageType;
        registry.unregister(MODULE.ID, key); 
    });

    console.log("Unregistered damage types:", damageTypesToUnregister.map(dt => dt.key));
}

function reRegisterDamageTypes(damageTypesToReRegister) {
    const registry = pf1.registry.damageTypes;

    damageTypesToReRegister.forEach(damageType => {
        const { key, value } = damageType;
        registry.register(MODULE.ID, key, value); 
    });

    console.log("Re-registered damage types:", damageTypesToReRegister.map(dt => dt.key));
}

export const moduleConfig = {
    damageConfig,
    registerSettings,
    registerNasSettings,
    handleReadyHook,
    handleRegistryHook,
    handleSetupHook,
    syncWeaponDamageTypes,
    populateDefaultTypes
};
let _nasSettingsRegistered = false;

export function registerNasSettings() {
  if (_nasSettingsRegistered) return;
  _nasSettingsRegistered = true;

  game.settings.register(MODULE.ID, 'reorderAllConditions', {
    name: "Reorder All Conditions Alphabetically",
    hint: 'Toggle to reorder all conditions alphabetically or only new conditions added by this module.',
    scope: "world",
    config: true,
    type: Boolean,
    default: false,
  });

  game.settings.register(MODULE.ID, 'automaticBuffs', {
    name: 'Enable Automatic Buffs',
    hint: 'When enabled, the module will attempt to find and apply matching buffs when spells or consumables are used.',
    scope: 'world',
    config: true,
    type: Boolean,
    default: true
  });
  
  game.settings.register(MODULE.ID, 'buffAutomationMode', {
    name: 'Buff Automation Mode',
    hint: 'Choose how strict the buff automation should be when no targets are selected.',
    scope: 'world',
    config: true,
    type: String,
    choices: {
      "strict": "Strict (Block actions without targets)",
      "lenient": "Lenient (Allow, but notify if no targets)"
    },
    default: "strict"
  });
  
  game.settings.register(MODULE.ID, 'buffTargetFiltering', {
    name: 'Buff Target Filtering',
    hint: 'Choose how buff targets are filtered. "By Disposition" only applies buffs to targets with the same disposition as the caster, "All Targets" applies buffs to all selected targets, and "Manual Selection" prompts you to choose which targets receive the buff.',
    scope: 'world',
    config: true,
    type: String,
    choices: {
      "byDisposition": "By Disposition (Only same disposition)",
      "allTargets": "All Targets (No filtering)",
      "manualSelection": "Manual Selection (Choose targets)"
    },
    default: "byDisposition"
  });
  
  game.settings.registerMenu(MODULE.ID, 'buffCompendiaSelector', {
    name: 'Select Buff Compendia',
    label: 'Select Compendia',
    hint: 'Choose which compendia to include when searching for buff items.',
    icon: 'fas fa-book',
    type: BuffCompendiaSelector,
    restricted: true
  });

  const defaultCompendia = ["pf1.buffs"];
  if (game.packs.get("pf-content.pf-buffs")) {
    defaultCompendia.push("pf-content.pf-buffs");
  }

  game.settings.register(MODULE.ID, 'customBuffCompendia', {
    name: 'Custom Buff Compendia',
    hint: 'Select additional compendia containing buffs to include in the automated buff search.',
    scope: 'world',
    config: false,
    type: Array,
    default: defaultCompendia,
  });

  game.settings.register(MODULE.ID, 'handleConfused', {
    name: 'Automate Confused Condition Actions',
    hint: 'Enable to automatically generate a message at the start of each round to determine the actions of confused tokens.',
    scope: 'world',
    config: true,
    type: Boolean,
    default: true
  });

  game.settings.register(MODULE.ID, 'restrictMovement', {
    name: "Restrict Movement",
    hint: "Choose who is restricted from moving when affected by immobilizing conditions such as 'anchored', 'cowering', 'dazed', 'dying', 'helpless', 'paralyzed', 'petrified', or 'pinned'. 'Players Only' allows GMs to always move tokens. 'Disabled' will allow all movement.",
    scope: "world",
    config: true,
    type: String,
    choices: {
      "all": "All (GM and Players)",
      "players": "Players Only (GM can always move)",
      "disabled": "Disabled"
    },
    default: "disabled",
  });

  game.settings.register(MODULE.ID, 'autoApplyFF', {
    name: "Auto Apply Flat-Footed Condition",
    hint: "Enable to automatically apply the flat-footed condition to any token with an initiative roll result lower than the highest when combat begins.",
    scope: "world",
    config: true,
    type: Boolean,
    default: true,
  });

  game.settings.register(MODULE.ID, 'blindMovementCheck', {
    name: 'Enable Blind Movement Notification',
    hint: 'Enable to notify users to roll an Acrobatics check when a blind token attempts to move.',
    scope: 'world',
    config: true,
    type: Boolean,
    default: true
  });

  game.settings.register(MODULE.ID, 'disableAtZeroHP', {
    name: 'Apply Disabled Condition at 0 HP',
    hint: 'Automatically apply the disabled condition based on the selected option.',
    scope: 'world',
    config: true,
    type: String,
    choices: {
        "none": "No one",
        "npc": "NPC Only",
        "player": "Player Only",
        "everyone": "Everyone"
    },
    default: "everyone"
  });

  game.settings.register(MODULE.ID, 'autoApplyED', {
    name: 'Auto Apply Energy Drain',
    hint: 'Enable to automatically apply the energy drain condition to any token with negative levels.',
    scope: 'world',
    config: true,
    type: Boolean,
    default: true
  });

  game.settings.register(MODULE.ID, 'entangledGrappledHandling', {
    name: 'Concentration Check for Entangled and Grappled',
    hint: 'Choose whether to prompt concentration checks for entangled and/or grappled spellcasting.',
    scope: 'world',
    config: true,
    type: String,
    choices: {
      disabled: 'Disabled',
      grappled: 'Only Grappled',
      entangled: 'Only Entangled',
      both: 'Both'
    },
    default: 'disabled'
  });

  game.settings.register(MODULE.ID, 'grappledHandling', {
    name: 'Grappled Action Handling',
    hint: 'Choose how actions requiring two hands should be handled when grappled: Strict, Lenient, or Disabled.',
    scope: 'world',
    config: true,
    type: String,
    choices: {
      "disabled": "Disabled (No restrictions)",
      "strict": "Strict (Block two-handed actions)",
      "lenient": "Lenient (Allow with warning)"
    },
    default: "strict"
  });  
  
  game.settings.register(MODULE.ID, 'nauseatedHandling', {
    name: 'Nauseated Action Handling',
    hint: 'Choose how actions are handled when affected by the nauseated condition: Strict, Lenient, or Disabled.',
    scope: 'world',
    config: true,
    type: String,
    choices: {
      "disabled": "Disabled (No restrictions)",
      "strict": "Strict (Block non-move actions)",
      "lenient": "Lenient (Allow with warning)"
    },
    default: "strict"
  });
  
  game.settings.register(MODULE.ID, 'squeezingHandling', {
    name: 'Squeezing Action Handling',
    hint: 'Choose how actions are handled when affected by the squeezing condition: Strict, Lenient, or Disabled.',
    scope: 'world',
    config: true,
    type: String,
    choices: {
      "disabled": "Disabled (No restrictions)",
      "strict": "Strict (Block attack actions)",
      "lenient": "Lenient (Allow with warning)"
    },
    default: "strict"
  });  

  game.settings.register(MODULE.ID, 'unconsciousAtNegativeHP', {
      name: 'Apply Unconscious Condition at Negative HP',
      hint: 'Automatically apply the unconscious condition based on the selected option.',
      scope: 'world',
      config: true,
      type: String,
      choices: {
          "none": "No one",
          "npc": "NPC Only",
          "player": "Player Only",
          "everyone": "Everyone"
      },
      default: "everyone"
  });

  const isMonksCombatDetailsActive = game.modules.get('monks-combat-details')?.active;
  const monksAutoDefeatedSetting = isMonksCombatDetailsActive ? game.settings.get('monks-combat-details', 'auto-defeated') : 'none';
  const defaultApplyDeadCondition = monksAutoDefeatedSetting !== 'none' ? false : true;
  
  game.settings.register(MODULE.ID, 'applyDeadCondition', {
    name: 'Apply Dead Condition at Negative HP',
    hint: `Automatically apply the dead condition based on the selected option (players at negative max HP, NPCs at negative HP).${isMonksCombatDetailsActive ? ' Enabling this option will disable the Monks Combat Details auto defeated setting.' : ''}`,
    scope: 'world',
    config: true,
    type: String,
    choices: {
        "none": "No one",
        "npc": "NPC Only",
        "player": "Player Only",
        "player-negative-con-npc-negative-hp": "Player (Negative Max), NPC (Negative HP)",
        "everyone": "Everyone"
    },
    default: defaultApplyDeadCondition ? "everyone" : "none",
    onChange: async (value) => {
      if (value !== "none" && isMonksCombatDetailsActive) {
        const choice = await Dialog.confirm({
            title: "Conflict with Monks Combat Details",
            content: "Enabling this setting will disable the auto-defeated setting of Monks Combat Details. Do you want to proceed?",
            yes: () => true,
            no: () => false,
            defaultYes: false
        });
        if (choice) {
            await game.settings.set('monks-combat-details', 'auto-defeated', 'none');
            ui.notifications.info("Monks Combat Details auto-defeated setting has been disabled.");
        } else {
            await game.settings.set(MODULE.ID, 'applyDeadCondition', 'none');
            ui.notifications.warn("Apply Dead Condition setting has been disabled.");
        };
      };
    }
  });

  game.settings.register(MODULE.ID, 'removeDeadCondition', {
    name: 'Auto Remove Dead Condition',
    hint: 'Choose when to automatically remove the dead condition.',
    scope: 'world',
    config: true,
    type: String,
    choices: {
      disabled: 'Disabled',
      aboveNegativeCon: 'When HP is above negative max HP',
      nonNegative: 'When HP is 0 or more'
    },
    default: 'disabled'
  });

  game.settings.registerMenu(MODULE.ID, 'modifierNameSettings', {
    name: 'Customize Buff/Spell Modifiers',
    label: 'Customize Modifiers',
    hint: 'Edit the display names for common buff/spell modifiers (e.g., Lesser, Greater, Mass, Communal, etc.)',
    icon: 'fas fa-pen',
    type: ModifierNameSettingsForm,
    restricted: true
  });

  game.settings.registerMenu(MODULE.ID, 'variantMappingManager', {
    name: 'Manage Variant Mappings',
    label: 'Manage Variant Mappings',
    hint: 'View and remove remembered variant assignments for spells/abilities.',
    icon: 'fas fa-list',
    type: VariantMappingManager,
    restricted: true
  });

  game.settings.register(MODULE.ID, 'modifierNames', {
    name: 'Buff/Spell Modifier Names',
    hint: 'Stores the custom names for buff/spell modifiers.',
    scope: 'world',
    config: false,
    type: Object,
    default: {
      lesser: 'Lesser',
      minor: 'Minor',
      improved: 'Improved',
      greater: 'Greater',
      major: 'Major',
      supreme: 'Supreme',
      mass: 'Mass',
      communal: 'Communal'
    }
  });

  game.settings.register(MODULE.ID, 'communalHandling', {
    name: 'Communal Spell Duration Handling',
    hint: 'Choose how communal spell durations are divided among targets: divide evenly (prompt if not possible), or always prompt the caster to divide.',
    scope: 'world',
    config: true,
    type: String,
    choices: {
      even: 'Divide Evenly (Prompt if impossible)',
      prompt: 'Always Prompt Caster'
    },
    default: 'even'
  });

  game.settings.register(MODULE.ID, 'personalTargeting', {
    name: 'Personal Spell Targeting',
    hint: 'Choose whether personal spells can target tokens other than the caster.',
    scope: 'world',
    config: true,
    type: String,
    choices: {
      allow: 'Allow targets other than the caster',
      deny: 'Deny targets that are not the caster'
    },
    default: 'deny'
  });

  game.settings.register(MODULE.ID, 'variantTargetCap', {
    name: 'Variant Target Cap Handling',
    hint: 'When a spell has a detected target cap (e.g., Blessing of Fervor), choose whether to enforce it (cancel if exceeded) or only show a warning hint.',
    scope: 'world',
    config: true,
    type: String,
    choices: {
      hint: 'Hint only (allow over-cap)',
      enforce: 'Enforce cap (cancel when exceeded)'
    },
    default: 'hint'
  });
  
  game.settings.register(MODULE.ID, 'pairedBuffMappings', {
    name: 'Paired Buff Variant Mappings',
    hint: 'Internal store for per-spell variant to group assignments.',
    scope: 'world',
    config: false,
    type: Object,
    default: {}
  });

  registerNasSettingsUi();
}

let _nasSettingsUiRegistered = false;
function registerNasSettingsUi() {
  if (_nasSettingsUiRegistered) return;
  _nasSettingsUiRegistered = true;

  Hooks.on('renderSettingsConfig', (app, html, data) => {
  const moduleId = MODULE.ID;
  const isJQ = typeof html?.find === "function";

  const tab = isJQ
    ? html.find(`section.tab[data-tab="${moduleId}"]`)
    : html.querySelector(`section.tab[data-tab="${moduleId}"]`);
  const tabEl = isJQ ? tab?.[0] : tab;
  if (!tabEl) return;

  function findFormGroup(selector) {
    if (!tabEl) return null;
    if (isJQ) {
      return tab.find(selector).closest(".form-group");
    }
    const el = tabEl.querySelector(selector);
    return el ? el.closest(".form-group") : null;
  }

  function asElement(maybeJq) {
    if (!maybeJq) return null;
    if (maybeJq instanceof HTMLElement) return maybeJq;
    if (isJQ && maybeJq[0] instanceof HTMLElement) return maybeJq[0];
    return null;
  }

  function getSettingRow(settingKey) {
    return findFormGroup(`*[name="${moduleId}.${settingKey}"]`);
  }

  function getMenuRow(menuKey) {
    return findFormGroup(`button[data-key="${moduleId}.${menuKey}"]`);
  }

  function createSection({ id, title, open }) {
    const details = document.createElement("details");
    details.classList.add("nas-settings-section");
    details.dataset.sectionId = id;
    if (open) details.open = true;

    const summary = document.createElement("summary");

    const titleSpan = document.createElement("span");
    titleSpan.classList.add("nas-settings-section-title");
    titleSpan.textContent = title;

    const caret = document.createElement("span");
    caret.classList.add("nas-settings-section-caret");
    caret.innerHTML = "&#9654;"; 

    summary.appendChild(titleSpan);
    summary.appendChild(caret);

    const body = document.createElement("div");
    body.classList.add("nas-settings-section-body");

    details.appendChild(summary);
    details.appendChild(body);
    return { details, body };
  }

  let container = tabEl.querySelector(":scope > .nas-settings-sections");
  if (!container) {
    container = document.createElement("div");
    container.classList.add("nas-settings-sections");
    tabEl.prepend(container);
  }

  const sections = [
    {
      id: "general",
      title: "General",
      open: true,
      rows: [
        () => getSettingRow("saveRollTokenInteraction"),
        () => getSettingRow("reorderAllConditions"),
        () => getMenuRow("migrationTool")
      ]
    },
    {
      id: "buff",
      title: "Buff Automation",
      open: true,
      rows: [
        () => getSettingRow("automaticBuffs"),
        () => getMenuRow("buffCompendiaSelector"),
        () => getMenuRow("modifierNameSettings"),
        () => getMenuRow("variantMappingManager"),
        () => getSettingRow("buffAutomationMode"),
        () => getSettingRow("buffTargetFiltering"),
        () => getSettingRow("communalHandling"),
        () => getSettingRow("personalTargeting"),
        () => getSettingRow("variantTargetCap")
      ]
    },
    {
      id: "damage",
      title: "Damage Automation",
      open: false,
      rows: [
        () => getSettingRow("massiveDamage"),
        () => getMenuRow("damageTypePriorityMenu"),
        () => getMenuRow("customSetting")
      ]
    },
    {
      id: "condition",
      title: "Condition Automation",
      open: false,
      rows: [
        () => getSettingRow("handleConfused"),
        () => getSettingRow("restrictMovement"),
        () => getSettingRow("autoApplyFF"),
        () => getSettingRow("blindMovementCheck"),
        () => getSettingRow("disableAtZeroHP"),
        () => getSettingRow("autoApplyED"),
        () => getSettingRow("entangledGrappledHandling"),
        () => getSettingRow("grappledHandling"),
        () => getSettingRow("nauseatedHandling"),
        () => getSettingRow("squeezingHandling"),
        () => getSettingRow("unconsciousAtNegativeHP"),
        () => getSettingRow("applyDeadCondition"),
        () => getSettingRow("removeDeadCondition")
      ]
    },
    {
      id: "metamagic",
      title: "Metamagic Automation",
      open: false,
      rows: [
        () => getSettingRow("metamagicCastTimeRule"),
        () => getSettingRow("persistentSpellTargetMode")
      ]
    },
    {
      id: "translations",
      title: "Translations",
      open: false,
      rows: [() => getMenuRow("translationMenu")]
    }
  ];

  const existingSections = new Map(
    Array.from(container.querySelectorAll(".nas-settings-section")).map(el => [el.dataset.sectionId, el])
  );

  for (const spec of sections) {
    let details = existingSections.get(spec.id);
    let body;
    if (!details) {
      const created = createSection(spec);
      details = created.details;
      body = created.body;
      container.appendChild(details);
    } else {
      body = details.querySelector(":scope > .nas-settings-section-body");
    }

    for (const getRow of spec.rows) {
      const row = getRow();
      if (!row) continue;
      const rowEl = asElement(row);
      if (!rowEl) continue;
      body.appendChild(rowEl);
    }
  }

  const automaticBuffsRow = getSettingRow("automaticBuffs");
  const buffSelectorRow = getMenuRow("buffCompendiaSelector");
  const modifierNameSettingsRow = getMenuRow("modifierNameSettings");
  const variantMappingManagerRow = getMenuRow("variantMappingManager");
  const buffAutomationModeRow = getSettingRow("buffAutomationMode");
  const buffTargetFilteringRow = getSettingRow("buffTargetFiltering");
  const communalHandlingRow = getSettingRow("communalHandling");
  const personalTargetingRow = getSettingRow("personalTargeting");
  const variantTargetCapRow = getSettingRow("variantTargetCap");

  let automaticBuffsCheckbox;
  if (isJQ) {
    automaticBuffsCheckbox = automaticBuffsRow?.find?.("input");
  } else {
    automaticBuffsCheckbox = asElement(automaticBuffsRow)?.querySelector?.("input");
  }

  function toggleBuffSettingsVisibility(show, elements) {
    elements.forEach(element => {
      if (!element) return;
      if (isJQ) {
        if (show) element.show();
        else element.hide();
      } else {
        const el = asElement(element);
        if (!el) return;
        el.style.display = show ? "" : "none";
      }
    });
  }

  const dependentRows = [
    buffSelectorRow,
    modifierNameSettingsRow,
    variantMappingManagerRow,
    buffAutomationModeRow,
    buffTargetFilteringRow,
    communalHandlingRow,
    personalTargetingRow,
    variantTargetCapRow
  ];

  const isEnabled = automaticBuffsCheckbox
    ? isJQ
      ? automaticBuffsCheckbox.prop("checked")
      : automaticBuffsCheckbox.checked
    : false;
  toggleBuffSettingsVisibility(isEnabled, dependentRows);

  if (automaticBuffsCheckbox) {
    if (isJQ) {
      automaticBuffsCheckbox.off("change.nas").on("change.nas", function () {
        toggleBuffSettingsVisibility($(this).prop("checked"), dependentRows);
      });
    } else if (!automaticBuffsCheckbox.dataset.nasListenerAttached) {
      automaticBuffsCheckbox.dataset.nasListenerAttached = "true";
      automaticBuffsCheckbox.addEventListener("change", function () {
        toggleBuffSettingsVisibility(this.checked, dependentRows);
      });
    }
  }
  });
}






