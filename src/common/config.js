import { MODULE } from "./module.js";
import { checkboxChecked, elementFromHtmlLike, insertNasSettingsSectionsContainer, setElementVisible } from "./foundryCompat.js";
import { DamagePriorityForm } from "./settings/damagePriorityForm.js";
import { DamageTypeFormApplication } from "./settings/damageTypeForm.js";
import { TranslationForm } from "./settings/translationForm.js";
import { BuffCompendiaSelector } from "./settings/buffCompendiaSelector.js";
import { ModifierNameSettingsForm } from "./settings/modifierNameSettingsForm.js";
import { VariantMappingManager } from "./settings/variantMappingManager.js";
import { WoundDamageTypesForm } from "./settings/woundDamageTypesForm.js";
import { SqueezingAutomationConfigForm } from "./settings/squeezingAutomationConfigForm.js";
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

    game.settings.register(MODULE.ID, "metamagicChatCardNameMode", {
        name: game.i18n.localize("NAS.settings.metamagicChatCardNameMode.name"),
        hint: game.i18n.localize("NAS.settings.metamagicChatCardNameMode.hint"),
        scope: "world",
        config: true,
        type: String,
        choices: {
            off: game.i18n.localize("NAS.settings.metamagicChatCardNameMode.choices.off"),
            highest: game.i18n.localize("NAS.settings.metamagicChatCardNameMode.choices.highest")
        },
        default: "highest"
    });

    game.settings.register(MODULE.ID, "metamagicPreviewMode", {
        name: game.i18n.localize("NAS.settings.metamagicPreviewMode.name"),
        hint: game.i18n.localize("NAS.settings.metamagicPreviewMode.hint"),
        scope: "world",
        config: true,
        type: String,
        choices: {
            none: game.i18n.localize("NAS.common.labels.none"),
            concise: game.i18n.localize("NAS.settings.metamagicPreviewMode.choices.concise"),
            detailed: game.i18n.localize("NAS.settings.metamagicPreviewMode.choices.detailed")
        },
        default: "concise"
    });

    game.settings.register(MODULE.ID, "saveRollTokenInteraction", {
        name: game.i18n.localize("NAS.settings.saveRollTokenInteraction.name"),
        hint: game.i18n.localize("NAS.settings.saveRollTokenInteraction.hint"),
        scope: "world",
        config: true,
        type: Boolean,
        default: true
    });

    game.settings.register(MODULE.ID, "enhancedDiceTooltipMode", {
        name: game.i18n.localize("NAS.settings.enhancedDiceTooltipMode.name"),
        hint: game.i18n.localize("NAS.settings.enhancedDiceTooltipMode.hint"),
        scope: "world",
        config: true,
        type: String,
        choices: {
            off: game.i18n.localize("NAS.settings.enhancedDiceTooltipMode.choices.off"),
            labeled: game.i18n.localize("NAS.settings.enhancedDiceTooltipMode.choices.labeled"),
            detailed: game.i18n.localize("NAS.settings.enhancedDiceTooltipMode.choices.detailed")
        },
        default: "labeled"
    });

  game.settings.register(MODULE.ID, "enhancedCombatText", {
    name: game.i18n.localize("NAS.settings.enhancedCombatText.name"),
    hint: game.i18n.localize("NAS.settings.enhancedCombatText.hint"),
    scope: "world",
    config: false,
    type: Boolean,
    default: false
  });

  game.settings.register(MODULE.ID, "combatTextMode", {
    name: game.i18n.localize("NAS.settings.combatTextMode.name"),
    hint: game.i18n.localize("NAS.settings.combatTextMode.hint"),
    scope: "world",
    config: true,
    type: String,
    choices: {
      off: game.i18n.localize("NAS.settings.combatTextMode.choices.off"),
      enhanced: game.i18n.localize("NAS.settings.combatTextMode.choices.enhanced"),
      cinematic: game.i18n.localize("NAS.settings.combatTextMode.choices.cinematic")
    },
    default: "off"
  });

  game.settings.register(MODULE.ID, "cinematicCombatTextPreset", {
    name: game.i18n.localize("NAS.settings.cinematicCombatTextPreset.name"),
    hint: game.i18n.localize("NAS.settings.cinematicCombatTextPreset.hint"),
    scope: "world",
    config: true,
    type: String,
    choices: {
      arcLanes: game.i18n.localize("NAS.settings.cinematicCombatTextPreset.choices.arcLanes"),
      tokenSides: game.i18n.localize("NAS.settings.cinematicCombatTextPreset.choices.tokenSides")
    },
    default: "arcLanes"
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
        label: game.i18n.localize("NAS.common.buttons.configure"),
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
        name: game.i18n.localize("NAS.settings.translations.name"),
        scope: "world",
        config: false,
        type: Object,
        default: {
            hardness: "",
            construct: "",
            undead: "",
            constructClassNames: "",
            undeadClassNames: "",
            constructRaceNames: "",
            undeadRaceNames: ""
        }
    });

    game.settings.register(MODULE.ID, "metamagicNameCache", {
        scope: "world",
        config: false,
        type: Object,
        default: {}
    });

    game.settings.registerMenu(MODULE.ID, "translationMenu", {
        name: game.i18n.localize("NAS.settings.translationMenu.name"),
        label: game.i18n.localize("NAS.settings.translationMenu.label"),
        hint: game.i18n.localize("NAS.settings.translationMenu.hint"),
        icon: "fas fa-language",
        type: TranslationForm,
        restricted: true
    });
    
    if (!game.settings.settings.has(`${MODULE.ID}.migrationVersion`)) {
        game.settings.register(MODULE.ID, "migrationVersion", {
            name: game.i18n.localize("NAS.settings.migrationVersion.name"),
            scope: "world",
            config: false,
            type: String,
            default: ""
        });
    }

    if (!game.settings.settings.has(`${MODULE.ID}.migrationTool`)) {
        game.settings.registerMenu(MODULE.ID, "migrationTool", {
            name: game.i18n.localize("NAS.settings.migrationTool.name"),
            label: game.i18n.localize("NAS.settings.migrationTool.label"),
            hint: game.i18n.localize("NAS.settings.migrationTool.hint"),
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
            ui.notifications.info(game.i18n.localize("NAS.settings.config.migrationStartInfo"));

            unregisterDamageTypes(damageTypesToReRegister);

            await game.settings.set(MODULE.ID, "customDamageTypes", customDamageTypes);

            reRegisterDamageTypes(damageTypesToReRegister);

            ui.notifications.info(game.i18n.localize("NAS.settings.config.migrationCompleteInfo"));
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
    name: game.i18n.localize("NAS.settings.reorderAllConditions.name"),
    hint: game.i18n.localize("NAS.settings.reorderAllConditions.hint"),
    scope: "world",
    config: true,
    type: Boolean,
    default: false,
  });

  game.settings.register(MODULE.ID, 'enableDamageAutomation', {
    name: game.i18n.localize("NAS.settings.enableDamageAutomation.name"),
    hint: game.i18n.localize("NAS.settings.enableDamageAutomation.hint"),
    scope: "world",
    config: true,
    type: Boolean,
    default: true,
  });

  game.settings.register(MODULE.ID, 'enableWoundsVigorAutomation', {
    name: game.i18n.localize("NAS.settings.enableWoundsVigorAutomation.name"),
    hint: game.i18n.localize("NAS.settings.enableWoundsVigorAutomation.hint"),
    scope: "world",
    config: true,
    type: Boolean,
    default: false,
  });

  game.settings.register(MODULE.ID, 'woundsVigorActionTaxMode', {
    name: game.i18n.localize("NAS.settings.woundsVigorActionTaxMode.name"),
    hint: game.i18n.localize("NAS.settings.woundsVigorActionTaxMode.hint"),
    scope: "world",
    config: true,
    type: String,
    choices: {
      strict: game.i18n.localize("NAS.settings.woundsVigorActionTaxMode.choices.strict"),
      broad: game.i18n.localize("NAS.settings.woundsVigorActionTaxMode.choices.broad")
    },
    default: "strict",
  });

  game.settings.register(MODULE.ID, 'woundsVigorNoWoundsConstructUndead', {
    name: game.i18n.localize("NAS.settings.woundsVigorNoWoundsConstructUndead.name"),
    hint: game.i18n.localize("NAS.settings.woundsVigorNoWoundsConstructUndead.hint"),
    scope: "world",
    config: true,
    type: Boolean,
    default: false,
  });

  game.settings.register(MODULE.ID, "woundsVigorWoundDamageTypeIds", {
    scope: "world",
    config: false,
    type: Array,
    default: ["negative", "positive"]
  });

  game.settings.registerMenu(MODULE.ID, "woundsVigorWoundDamageTypeMenu", {
    name: game.i18n.localize("NAS.forms.woundDamageTypesForm.title"),
    label: game.i18n.localize("NAS.common.buttons.configure"),
    hint: game.i18n.localize("NAS.settings.woundsVigorWoundDamageTypeMenu.hint"),
    icon: "fas fa-heart-pulse",
    type: WoundDamageTypesForm,
    restricted: true
  });

  game.settings.register(MODULE.ID, 'enableMetamagicAutomation', {
    name: game.i18n.localize("NAS.settings.enableMetamagicAutomation.name"),
    hint: game.i18n.localize("NAS.settings.enableMetamagicAutomation.hint"),
    scope: "world",
    config: true,
    type: Boolean,
    default: true,
  });

  game.settings.register(MODULE.ID, "enforceSpellAbilityMinimum", {
    name: game.i18n.localize("NAS.settings.enforceSpellAbilityMinimum.name"),
    hint: game.i18n.localize("NAS.settings.enforceSpellAbilityMinimum.hint"),
    scope: "world",
    config: true,
    type: Boolean,
    default: false,
  });

  game.settings.register(MODULE.ID, 'automaticBuffs', {
    name: game.i18n.localize("NAS.settings.automaticBuffs.name"),
    hint: game.i18n.localize("NAS.settings.automaticBuffs.hint"),
    scope: 'world',
    config: true,
    type: Boolean,
    default: true
  });
  
  game.settings.register(MODULE.ID, 'buffAutomationMode', {
    name: game.i18n.localize("NAS.settings.buffAutomationMode.name"),
    hint: game.i18n.localize("NAS.settings.buffAutomationMode.hint"),
    scope: 'world',
    config: true,
    type: String,
    choices: {
      "strict": game.i18n.localize("NAS.settings.buffAutomationMode.choices.strict"),
      "lenient": game.i18n.localize("NAS.settings.buffAutomationMode.choices.lenient")
    },
    default: "strict"
  });
  
  game.settings.register(MODULE.ID, 'buffTargetFiltering', {
    name: game.i18n.localize("NAS.settings.buffTargetFiltering.name"),
    hint: game.i18n.localize("NAS.settings.buffTargetFiltering.hint"),
    scope: 'world',
    config: true,
    type: String,
    choices: {
      "byDisposition": game.i18n.localize("NAS.settings.buffTargetFiltering.choices.disposition"),
      "allTargets": game.i18n.localize("NAS.settings.buffTargetFiltering.choices.all"),
      "manualSelection": game.i18n.localize("NAS.settings.buffTargetFiltering.choices.manual")
    },
    default: "byDisposition"
  });
  
  game.settings.registerMenu(MODULE.ID, 'buffCompendiaSelector', {
    name: game.i18n.localize("NAS.forms.buffCompendiaSelector.title"),
    label: game.i18n.localize("NAS.settings.buffCompendiaSelector.label"),
    hint: game.i18n.localize("NAS.settings.buffCompendiaSelector.hint"),
    icon: 'fas fa-book',
    type: BuffCompendiaSelector,
    restricted: true
  });

  const defaultCompendia = ["pf1.buffs", `${MODULE.ID}.Buffs`];
  if (game.packs.get("pf-content.pf-buffs")) {
    defaultCompendia.push("pf-content.pf-buffs");
  }

  game.settings.register(MODULE.ID, 'customBuffCompendia', {
    name: game.i18n.localize("NAS.settings.customBuffCompendia.name"),
    hint: game.i18n.localize("NAS.settings.customBuffCompendia.hint"),
    scope: 'world',
    config: false,
    type: Array,
    default: defaultCompendia,
  });

  game.settings.register(MODULE.ID, 'automateConditions', {
    name: game.i18n.localize("NAS.settings.automateConditions.name"),
    hint: game.i18n.localize("NAS.settings.automateConditions.hint"),
    scope: 'world',
    config: true,
    type: Boolean,
    default: true
  });

  game.settings.register(MODULE.ID, 'handleConfused', {
    name: game.i18n.localize("NAS.settings.handleConfused.name"),
    hint: game.i18n.localize("NAS.settings.handleConfused.hint"),
    scope: 'world',
    config: true,
    type: Boolean,
    default: true
  });

  game.settings.register(MODULE.ID, 'restrictMovement', {
    name: game.i18n.localize("NAS.settings.restrictMovement.name"),
    hint: game.i18n.localize("NAS.settings.restrictMovement.hint"),
    scope: "world",
    config: true,
    type: String,
    choices: {
        "all": game.i18n.localize("NAS.settings.restrictMovement.choices.all"),
        "players": game.i18n.localize("NAS.settings.restrictMovement.choices.players"),
        "disabled": game.i18n.localize("NAS.common.choices.handling.disabled")
    },
    default: "disabled",
  });

  game.settings.register(MODULE.ID, 'autoApplyFF', {
    name: game.i18n.localize("NAS.settings.autoApplyFF.name"),
    hint: game.i18n.localize("NAS.settings.autoApplyFF.hint"),
    scope: "world",
    config: true,
    type: Boolean,
    default: true,
  });

  game.settings.register(MODULE.ID, 'skipSurprisedTokens', {
    name: game.i18n.localize("NAS.settings.skipSurprisedTokens.name"),
    hint: game.i18n.localize("NAS.settings.skipSurprisedTokens.hint"),
    scope: "world",
    config: true,
    type: Boolean,
    default: false,
  });

  game.settings.register(MODULE.ID, 'blindMovementCheck', {
    name: game.i18n.localize("NAS.settings.blindMovementCheck.name"),
    hint: game.i18n.localize("NAS.settings.blindMovementCheck.hint"),
    scope: 'world',
    config: true,
    type: Boolean,
    default: true
  });

  game.settings.register(MODULE.ID, 'disableAtZeroHP', {
    name: game.i18n.localize("NAS.settings.disableAtZeroHP.name"),
    hint: game.i18n.localize("NAS.settings.disableAtZeroHP.hint"),
    scope: 'world',
    config: true,
    type: String,
    choices: {
        "none": game.i18n.localize("NAS.common.choices.audience.none"),
        "npc": game.i18n.localize("NAS.common.choices.audience.npc"),
        "player": game.i18n.localize("NAS.common.choices.audience.player"),
        "everyone": game.i18n.localize("NAS.common.choices.audience.everyone")
    },
    default: "everyone"
  });

  game.settings.register(MODULE.ID, 'autoApplyED', {
    name: game.i18n.localize("NAS.settings.autoApplyED.name"),
    hint: game.i18n.localize("NAS.settings.autoApplyED.hint"),
    scope: 'world',
    config: true,
    type: Boolean,
    default: true
  });

  game.settings.register(MODULE.ID, 'entangledGrappledHandling', {
    name: game.i18n.localize("NAS.settings.entangledGrappledHandling.name"),
    hint: game.i18n.localize("NAS.settings.entangledGrappledHandling.hint"),
    scope: 'world',
    config: true,
    type: String,
    choices: {
      disabled: game.i18n.localize("NAS.common.choices.handling.disabled"),
      grappled: game.i18n.localize("NAS.settings.entangledGrappledHandling.choices.grappled"),
      entangled: game.i18n.localize("NAS.settings.entangledGrappledHandling.choices.entangled"),
      both: game.i18n.localize("NAS.settings.entangledGrappledHandling.choices.both")
    },
    default: 'disabled'
  });

  game.settings.register(MODULE.ID, 'grappledHandling', {
    name: game.i18n.localize("NAS.settings.grappledHandling.name"),
    hint: game.i18n.localize("NAS.settings.grappledHandling.hint"),
    scope: 'world',
    config: true,
    type: String,
    choices: {
      "disabled": game.i18n.localize("NAS.common.choices.handling.disabledNoRestrictions"),
      "strict": game.i18n.localize("NAS.settings.grappledHandling.choices.strict"),
      "lenient": game.i18n.localize("NAS.common.choices.handling.lenientWarning")
    },
    default: "strict"
  });  
  
  game.settings.register(MODULE.ID, 'nauseatedHandling', {
    name: game.i18n.localize("NAS.settings.nauseatedHandling.name"),
    hint: game.i18n.localize("NAS.settings.nauseatedHandling.hint"),
    scope: 'world',
    config: true,
    type: String,
    choices: {
      "disabled": game.i18n.localize("NAS.common.choices.handling.disabledNoRestrictions"),
      "strict": game.i18n.localize("NAS.settings.nauseatedHandling.choices.strict"),
      "lenient": game.i18n.localize("NAS.common.choices.handling.lenientWarning")
    },
    default: "strict"
  });

  game.settings.register(MODULE.ID, 'automateSqueezing', {
    name: game.i18n.localize("NAS.settings.automateSqueezing.name"),
    hint: game.i18n.localize("NAS.settings.automateSqueezing.hint"),
    scope: 'world',
    config: true,
    type: Boolean,
    default: true
  });
  
  game.settings.register(MODULE.ID, 'squeezedHandling', {
    name: game.i18n.localize("NAS.settings.squeezedHandling.name"),
    hint: game.i18n.localize("NAS.settings.squeezedHandling.hint"),
    scope: 'world',
    config: false,
    type: String,
    choices: {
      "disabled": game.i18n.localize("NAS.common.choices.handling.disabledNoRestrictions"),
      "strict": game.i18n.localize("NAS.settings.squeezedHandling.choices.strict"),
      "lenient": game.i18n.localize("NAS.common.choices.handling.lenientWarning")
    },
    default: "strict"
  });  

  game.settings.register(MODULE.ID, 'squeezedExitHandling', {
    name: game.i18n.localize("NAS.settings.squeezedExitHandling.name"),
    hint: game.i18n.localize("NAS.settings.squeezedExitHandling.hint"),
    scope: 'world',
    config: false,
    type: String,
    choices: {
      "count": game.i18n.localize("NAS.settings.squeezedExitHandling.choices.count"),
      "ignore": game.i18n.localize("NAS.settings.squeezedExitHandling.choices.ignore")
    },
    default: "count"
  });

  game.settings.register(MODULE.ID, 'squeezedEscapeFailureHandling', {
    name: game.i18n.localize("NAS.settings.squeezedEscapeFailureHandling.name"),
    hint: game.i18n.localize("NAS.settings.squeezedEscapeFailureHandling.hint"),
    scope: 'world',
    config: false,
    type: String,
    choices: {
      "stopBeforeNarrow": game.i18n.localize("NAS.settings.squeezedEscapeFailureHandling.choices.stopBeforeNarrow"),
      "enterFirstNarrowSquare": game.i18n.localize("NAS.settings.squeezedEscapeFailureHandling.choices.enterFirstNarrowSquare")
    },
    default: "stopBeforeNarrow"
  });

  game.settings.register(MODULE.ID, 'squeezingMediumBodyWidth', {
    name: game.i18n.localize("NAS.settings.squeezingMediumBodyWidth.name"),
    hint: game.i18n.localize("NAS.settings.squeezingMediumBodyWidth.hint"),
    scope: 'world',
    config: false,
    type: Number,
    default: 0.75
  });

  game.settings.register(MODULE.ID, 'squeezingMediumHeadWidth', {
    name: game.i18n.localize("NAS.settings.squeezingMediumHeadWidth.name"),
    hint: game.i18n.localize("NAS.settings.squeezingMediumHeadWidth.hint"),
    scope: 'world',
    config: false,
    type: Number,
    default: 0.25
  });

  game.settings.register(MODULE.ID, 'squeezingEscapeArtistDC', {
    name: game.i18n.localize("NAS.settings.squeezingEscapeArtistDC.name"),
    hint: game.i18n.localize("NAS.settings.squeezingEscapeArtistDC.hint"),
    scope: 'world',
    config: false,
    type: Number,
    default: 30
  });

  game.settings.registerMenu(MODULE.ID, 'squeezingAutomationConfig', {
    name: game.i18n.localize("NAS.forms.squeezingAutomationConfig.title"),
    label: game.i18n.localize("NAS.common.buttons.configure"),
    hint: game.i18n.localize("NAS.settings.squeezingAutomationConfig.hint"),
    icon: 'fas fa-arrows-left-right',
    type: SqueezingAutomationConfigForm,
    restricted: true
  });

  game.settings.register(MODULE.ID, 'unconsciousAtNegativeHP', {
      name: game.i18n.localize("NAS.settings.unconsciousAtNegativeHP.name"),
      hint: game.i18n.localize("NAS.settings.unconsciousAtNegativeHP.hint"),
      scope: 'world',
      config: true,
      type: String,
      choices: {
          "none": game.i18n.localize("NAS.common.choices.audience.none"),
          "npc": game.i18n.localize("NAS.common.choices.audience.npc"),
          "player": game.i18n.localize("NAS.common.choices.audience.player"),
          "everyone": game.i18n.localize("NAS.common.choices.audience.everyone")
      },
      default: "everyone"
  });

  game.settings.register(MODULE.ID, 'enableDyingAutomation', {
    name: game.i18n.localize("NAS.settings.enableDyingAutomation.name"),
    hint: game.i18n.localize("NAS.settings.enableDyingAutomation.hint"),
    scope: 'world',
    config: true,
    type: Boolean,
    default: false
  });

  game.settings.register(MODULE.ID, 'dyingHandling', {
    name: game.i18n.localize("NAS.settings.dyingHandling.name"),
    hint: game.i18n.localize("NAS.settings.dyingHandling.hint"),
    scope: 'world',
    config: true,
    type: String,
    choices: {
      "disabled": game.i18n.localize("NAS.common.choices.handling.disabledNoRestrictions"),
      "strict": game.i18n.localize("NAS.settings.dyingHandling.choices.strict"),
      "lenient": game.i18n.localize("NAS.common.choices.handling.lenientWarning")
    },
    default: "strict"
  });

  const isMonksCombatDetailsActive = game.modules.get('monks-combat-details')?.active;
  const monksAutoDefeatedSetting = isMonksCombatDetailsActive ? game.settings.get('monks-combat-details', 'auto-defeated') : 'none';
  const defaultApplyDeadCondition = monksAutoDefeatedSetting !== 'none' ? false : true;
  
  game.settings.register(MODULE.ID, 'applyDeadCondition', {
    name: game.i18n.localize("NAS.settings.applyDeadCondition.name"),
    hint: game.i18n.format("NAS.settings.applyDeadCondition.hint", {
      monksSuffix: isMonksCombatDetailsActive ? game.i18n.localize("NAS.settings.applyDeadCondition.monksSuffix") : ""
    }),
    scope: 'world',
    config: true,
    type: String,
    choices: {
        "none": game.i18n.localize("NAS.common.choices.audience.none"),
        "npc": game.i18n.localize("NAS.common.choices.audience.npc"),
        "player": game.i18n.localize("NAS.common.choices.audience.player"),
        "player-negative-con-npc-negative-hp": game.i18n.localize("NAS.settings.applyDeadCondition.choices.player-negative-con-npc-negative-hp"),
        "everyone": game.i18n.localize("NAS.common.choices.audience.everyone")
    },
    default: defaultApplyDeadCondition ? "everyone" : "none",
    onChange: async (value) => {
      if (value !== "none" && isMonksCombatDetailsActive) {
        const choice = await Dialog.confirm({
            title: game.i18n.localize("NAS.settings.applyDeadCondition.conflictTitle"),
            content: game.i18n.localize("NAS.settings.applyDeadCondition.conflictContent"),
            yes: () => true,
            no: () => false,
            defaultYes: false
        });
        if (choice) {
            await game.settings.set('monks-combat-details', 'auto-defeated', 'none');
            ui.notifications.info(game.i18n.localize("NAS.settings.applyDeadCondition.monksDisabledInfo"));
        } else {
            await game.settings.set(MODULE.ID, 'applyDeadCondition', 'none');
            ui.notifications.warn(game.i18n.localize("NAS.settings.applyDeadCondition.applyDeadDisabledWarning"));
        };
      };
    }
  });

  game.settings.register(MODULE.ID, 'removeDeadCondition', {
    name: game.i18n.localize("NAS.settings.removeDeadCondition.name"),
    hint: game.i18n.localize("NAS.settings.removeDeadCondition.hint"),
    scope: 'world',
    config: true,
    type: String,
    choices: {
      disabled: game.i18n.localize("NAS.common.choices.handling.disabled"),
      aboveNegativeCon: game.i18n.localize("NAS.settings.removeDeadCondition.choices.aboveNegativeCon"),
      nonNegative: game.i18n.localize("NAS.settings.removeDeadCondition.choices.nonNegative")
    },
    default: 'disabled'
  });

  game.settings.registerMenu(MODULE.ID, 'modifierNameSettings', {
    name: game.i18n.localize("NAS.forms.modifierNames.title"),
    label: game.i18n.localize("NAS.settings.modifierNameSettings.label"),
    hint: game.i18n.localize("NAS.settings.modifierNameSettings.hint"),
    icon: 'fas fa-pen',
    type: ModifierNameSettingsForm,
    restricted: true
  });

  game.settings.registerMenu(MODULE.ID, 'variantMappingManager', {
    name: game.i18n.localize("NAS.forms.variantMappingManager.title"),
    label: game.i18n.localize("NAS.forms.variantMappingManager.title"),
    hint: game.i18n.localize("NAS.settings.variantMappingManager.hint"),
    icon: 'fas fa-list',
    type: VariantMappingManager,
    restricted: true
  });

  game.settings.register(MODULE.ID, 'modifierNames', {
    name: game.i18n.localize("NAS.settings.modifierNames.name"),
    hint: game.i18n.localize("NAS.settings.modifierNames.hint"),
    scope: 'world',
    config: false,
    type: Object,
    default: {
      lesser: game.i18n.localize("NAS.forms.modifierNames.labels.lesser"),
      minor: game.i18n.localize("NAS.forms.modifierNames.labels.minor"),
      improved: game.i18n.localize("NAS.forms.modifierNames.labels.improved"),
      greater: game.i18n.localize("NAS.forms.modifierNames.labels.greater"),
      major: game.i18n.localize("NAS.forms.modifierNames.labels.major"),
      supreme: game.i18n.localize("NAS.forms.modifierNames.labels.supreme"),
      mass: game.i18n.localize("NAS.forms.modifierNames.labels.mass"),
      communal: game.i18n.localize("NAS.forms.modifierNames.labels.communal")
    }
  });

  game.settings.register(MODULE.ID, 'communalHandling', {
    name: game.i18n.localize("NAS.settings.communalHandling.name"),
    hint: game.i18n.localize("NAS.settings.communalHandling.hint"),
    scope: 'world',
    config: true,
    type: String,
    choices: {
      even: game.i18n.localize("NAS.settings.communalHandling.choices.even"),
      prompt: game.i18n.localize("NAS.settings.communalHandling.choices.prompt")
    },
    default: 'even'
  });

  game.settings.register(MODULE.ID, 'personalTargeting', {
    name: game.i18n.localize("NAS.settings.personalTargeting.name"),
    hint: game.i18n.localize("NAS.settings.personalTargeting.hint"),
    scope: 'world',
    config: true,
    type: String,
    choices: {
      allow: game.i18n.localize("NAS.settings.personalTargeting.choices.allow"),
      deny: game.i18n.localize("NAS.settings.personalTargeting.choices.deny")
    },
    default: 'deny'
  });

  game.settings.register(MODULE.ID, 'variantTargetCap', {
    name: game.i18n.localize("NAS.settings.variantTargetCap.name"),
    hint: game.i18n.localize("NAS.settings.variantTargetCap.hint"),
    scope: 'world',
    config: true,
    type: String,
    choices: {
      hint: game.i18n.localize("NAS.settings.variantTargetCap.choices.warn"),
      enforce: game.i18n.localize("NAS.settings.variantTargetCap.choices.enforce")
    },
    default: 'hint'
  });
  
  game.settings.register(MODULE.ID, 'pairedBuffMappings', {
    name: game.i18n.localize("NAS.settings.pairedBuffMappings.name"),
    hint: game.i18n.localize("NAS.settings.pairedBuffMappings.hint"),
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

  function getSavedApplicationsScheme() {
    let applicationsScheme = "dark";
    try {
      const uiConfig = game.settings.get("core", "uiConfig");
      const scheme = uiConfig?.colorScheme?.applications;
      if (scheme === "light") applicationsScheme = "light";
    } catch (_err) {
    }
    return applicationsScheme;
  }

  function getSchemeFromUiConfigValue(value) {
    if (value === "light") return "light";
    if (value === "dark") return "dark";

    try {
      return globalThis.matchMedia?.("(prefers-color-scheme: light)")?.matches ? "light" : "dark";
    } catch (_err) {
      return "dark";
    }
  }

  function applyNasScheme(scheme) {
    document.querySelectorAll(".nas-settings-sections").forEach((el) => {
      el.dataset.nasScheme = scheme;
    });
  }

  Hooks.on("renderUIConfig", (app, html) => {
    const root = elementFromHtmlLike(html);
    const selector = 'select[name="core.uiConfig.colorScheme.applications"]';
    const sel = root?.querySelector?.(selector);
    if (!sel) return;

    if (!sel.dataset.nasListenerAttached) {
      sel.dataset.nasListenerAttached = "true";
      sel.addEventListener("change", () => {
        applyNasScheme(getSchemeFromUiConfigValue(sel.value));
      });
    }

    applyNasScheme(getSchemeFromUiConfigValue(sel.value));
  });

  Hooks.on("closeUIConfig", () => {
    applyNasScheme(getSavedApplicationsScheme());
  });

  Hooks.on('renderSettingsConfig', (app, html, data) => {
  const moduleId = MODULE.ID;
  const root = elementFromHtmlLike(html);

  const tabEl = root?.querySelector?.(`section.tab[data-tab="${moduleId}"]`);
  if (!tabEl) return;

  function findFormGroup(selector) {
    if (!tabEl) return null;
    const el = tabEl.querySelector(selector);
    return el ? el.closest(".form-group") : null;
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
    insertNasSettingsSectionsContainer(tabEl, container);
  }

  container.dataset.nasScheme = getSavedApplicationsScheme();

  const sections = [
    {
      id: "general",
      title: game.i18n.localize("NAS.settings.sections.general"),
      open: false,
      rows: [
        () => getSettingRow("saveRollTokenInteraction"),
        () => getSettingRow("enhancedDiceTooltipMode"),
        () => getSettingRow("combatTextMode"),
        () => getSettingRow("cinematicCombatTextPreset"),
        () => getSettingRow("enforceSpellAbilityMinimum"),
        () => getSettingRow("reorderAllConditions"),
        () => getMenuRow("migrationTool")
      ]
    },
    {
      id: "buff",
      title: game.i18n.localize("NAS.settings.sections.buff"),
      open: false,
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
      title: game.i18n.localize("NAS.settings.sections.damage"),
      open: false,
      rows: [
        () => getSettingRow("enableDamageAutomation"),
        () => getSettingRow("enableWoundsVigorAutomation"),
        () => getSettingRow("woundsVigorActionTaxMode"),
        () => getSettingRow("woundsVigorNoWoundsConstructUndead"),
        () => getMenuRow("woundsVigorWoundDamageTypeMenu"),
        () => getSettingRow("massiveDamage"),
        () => getMenuRow("damageTypePriorityMenu"),
        () => getMenuRow("customSetting")
      ]
    },
    {
      id: "condition",
      title: game.i18n.localize("NAS.settings.sections.condition"),
      open: false,
      rows: [
        () => getSettingRow("automateConditions"),
        () => getSettingRow("handleConfused"),
        () => getSettingRow("restrictMovement"),
        () => getSettingRow("autoApplyFF"),
        () => getSettingRow("skipSurprisedTokens"),
        () => getSettingRow("blindMovementCheck"),
        () => getSettingRow("disableAtZeroHP"),
        () => getSettingRow("autoApplyED"),
        () => getSettingRow("entangledGrappledHandling"),
        () => getSettingRow("grappledHandling"),
        () => getSettingRow("nauseatedHandling"),
        () => getSettingRow("automateSqueezing"),
        () => getMenuRow("squeezingAutomationConfig"),
        () => getSettingRow("unconsciousAtNegativeHP"),
        () => getSettingRow("enableDyingAutomation"),
        () => getSettingRow("dyingHandling"),
        () => getSettingRow("applyDeadCondition"),
        () => getSettingRow("removeDeadCondition")
      ]
    },
    {
      id: "metamagic",
      title: game.i18n.localize("NAS.settings.sections.metamagic"),
      open: false,
      rows: [
        () => getSettingRow("enableMetamagicAutomation"),
        () => getSettingRow("metamagicCastTimeRule"),
        () => getSettingRow("persistentSpellTargetMode"),
        () => getSettingRow("metamagicChatCardNameMode"),
        () => getSettingRow("metamagicPreviewMode")
      ]
    },
    {
      id: "translations",
      title: game.i18n.localize("NAS.settings.sections.translations"),
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
      body.appendChild(row);
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

  const automaticBuffsCheckbox = automaticBuffsRow?.querySelector?.("input");

  function toggleBuffSettingsVisibility(show, elements) {
    elements.forEach(element => {
      setElementVisible(element, show);
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

  const isEnabled = checkboxChecked(automaticBuffsCheckbox);
  toggleBuffSettingsVisibility(isEnabled, dependentRows);

  if (automaticBuffsCheckbox && !automaticBuffsCheckbox.dataset.nasListenerAttached) {
      automaticBuffsCheckbox.dataset.nasListenerAttached = "true";
      automaticBuffsCheckbox.addEventListener("change", function () {
        toggleBuffSettingsVisibility(this.checked, dependentRows);
      });
  }

  const enableDamageAutomationRow = getSettingRow("enableDamageAutomation");
  const enableWoundsVigorAutomationRow = getSettingRow("enableWoundsVigorAutomation");
  const woundsVigorActionTaxModeRow = getSettingRow("woundsVigorActionTaxMode");
  const woundsVigorNoWoundsConstructUndeadRow = getSettingRow("woundsVigorNoWoundsConstructUndead");
  const woundsVigorWoundDamageTypeMenuRow = getMenuRow("woundsVigorWoundDamageTypeMenu");
  const massiveDamageRow = getSettingRow("massiveDamage");
  const damageTypePriorityMenuRow = getMenuRow("damageTypePriorityMenu");

  const enableDamageAutomationCheckbox = enableDamageAutomationRow?.querySelector?.("input");

  const damageDependentRows = [
    enableWoundsVigorAutomationRow,
    woundsVigorActionTaxModeRow,
    woundsVigorNoWoundsConstructUndeadRow,
    woundsVigorWoundDamageTypeMenuRow,
    massiveDamageRow,
    damageTypePriorityMenuRow
  ];

  const damageEnabled = checkboxChecked(enableDamageAutomationCheckbox);
  toggleBuffSettingsVisibility(damageEnabled, damageDependentRows);

  if (enableDamageAutomationCheckbox && !enableDamageAutomationCheckbox.dataset.nasListenerAttached) {
      enableDamageAutomationCheckbox.dataset.nasListenerAttached = "true";
      enableDamageAutomationCheckbox.addEventListener("change", function () {
        toggleBuffSettingsVisibility(this.checked, damageDependentRows);
      });
  }

  const enableWoundsVigorAutomationCheckbox = enableWoundsVigorAutomationRow?.querySelector?.("input");

  const woundsVigorEnabled = checkboxChecked(enableWoundsVigorAutomationCheckbox);
  toggleBuffSettingsVisibility(woundsVigorEnabled, [woundsVigorActionTaxModeRow, woundsVigorNoWoundsConstructUndeadRow, woundsVigorWoundDamageTypeMenuRow]);

  if (enableWoundsVigorAutomationCheckbox && !enableWoundsVigorAutomationCheckbox.dataset.nasWvListenerAttached) {
      enableWoundsVigorAutomationCheckbox.dataset.nasWvListenerAttached = "true";
      enableWoundsVigorAutomationCheckbox.addEventListener("change", function () {
        toggleBuffSettingsVisibility(this.checked, [woundsVigorActionTaxModeRow, woundsVigorNoWoundsConstructUndeadRow, woundsVigorWoundDamageTypeMenuRow]);
      });
  }

  const enableDyingAutomationRow = getSettingRow("enableDyingAutomation");
  const dyingHandlingRow = getSettingRow("dyingHandling");
  const enableDyingAutomationCheckbox = enableDyingAutomationRow?.querySelector?.("input");

  const dyingEnabled = checkboxChecked(enableDyingAutomationCheckbox);
  toggleBuffSettingsVisibility(dyingEnabled, [dyingHandlingRow]);

  if (enableDyingAutomationCheckbox && !enableDyingAutomationCheckbox.dataset.nasDyingListenerAttached) {
      enableDyingAutomationCheckbox.dataset.nasDyingListenerAttached = "true";
      enableDyingAutomationCheckbox.addEventListener("change", function () {
        toggleBuffSettingsVisibility(this.checked, [dyingHandlingRow]);
      });
  }

  const automateSqueezingRow = getSettingRow("automateSqueezing");
  const squeezingAutomationConfigRow = getMenuRow("squeezingAutomationConfig");
  const automateSqueezingCheckbox = automateSqueezingRow?.querySelector?.("input");

  const squeezingAutomationEnabled = checkboxChecked(automateSqueezingCheckbox);
  toggleBuffSettingsVisibility(squeezingAutomationEnabled, [squeezingAutomationConfigRow]);

  if (automateSqueezingCheckbox && !automateSqueezingCheckbox.dataset.nasSqueezeListenerAttached) {
      automateSqueezingCheckbox.dataset.nasSqueezeListenerAttached = "true";
      automateSqueezingCheckbox.addEventListener("change", function () {
        toggleBuffSettingsVisibility(this.checked, [squeezingAutomationConfigRow]);
      });
  }

  const automateConditionsRow = getSettingRow("automateConditions");
  const handleConfusedRow = getSettingRow("handleConfused");
  const restrictMovementRow = getSettingRow("restrictMovement");
  const autoApplyFFRow = getSettingRow("autoApplyFF");
  const skipSurprisedTokensRow = getSettingRow("skipSurprisedTokens");
  const blindMovementCheckRow = getSettingRow("blindMovementCheck");
  const disableAtZeroHPRow = getSettingRow("disableAtZeroHP");
  const autoApplyEDRow = getSettingRow("autoApplyED");
  const entangledGrappledHandlingRow = getSettingRow("entangledGrappledHandling");
  const grappledHandlingRow = getSettingRow("grappledHandling");
  const nauseatedHandlingRow = getSettingRow("nauseatedHandling");
  const unconsciousAtNegativeHPRow = getSettingRow("unconsciousAtNegativeHP");
  const applyDeadConditionRow = getSettingRow("applyDeadCondition");
  const removeDeadConditionRow = getSettingRow("removeDeadCondition");

  const automateConditionsCheckbox = automateConditionsRow?.querySelector?.("input");

  const conditionDependentRows = [
    handleConfusedRow,
    restrictMovementRow,
    autoApplyFFRow,
    skipSurprisedTokensRow,
    blindMovementCheckRow,
    disableAtZeroHPRow,
    autoApplyEDRow,
    entangledGrappledHandlingRow,
    grappledHandlingRow,
    nauseatedHandlingRow,
    automateSqueezingRow,
    squeezingAutomationConfigRow,
    unconsciousAtNegativeHPRow,
    enableDyingAutomationRow,
    dyingHandlingRow,
    applyDeadConditionRow,
    removeDeadConditionRow
  ];

  const conditionAutomationEnabled = checkboxChecked(automateConditionsCheckbox);
  toggleBuffSettingsVisibility(conditionAutomationEnabled, conditionDependentRows);

  if (automateConditionsCheckbox && !automateConditionsCheckbox.dataset.nasConditionsListenerAttached) {
      automateConditionsCheckbox.dataset.nasConditionsListenerAttached = "true";
      automateConditionsCheckbox.addEventListener("change", function () {
        toggleBuffSettingsVisibility(this.checked, conditionDependentRows);
      });
  }

  const enableMetamagicAutomationRow = getSettingRow("enableMetamagicAutomation");
  const metamagicCastTimeRuleRow = getSettingRow("metamagicCastTimeRule");
  const persistentSpellTargetModeRow = getSettingRow("persistentSpellTargetMode");
  const metamagicChatCardNameModeRow = getSettingRow("metamagicChatCardNameMode");
  const metamagicPreviewModeRow = getSettingRow("metamagicPreviewMode");

  const enableMetamagicAutomationCheckbox = enableMetamagicAutomationRow?.querySelector?.("input");

  const metamagicDependentRows = [
    metamagicCastTimeRuleRow,
    persistentSpellTargetModeRow,
    metamagicChatCardNameModeRow,
    metamagicPreviewModeRow
  ];

  const metamagicEnabled = checkboxChecked(enableMetamagicAutomationCheckbox);
  toggleBuffSettingsVisibility(metamagicEnabled, metamagicDependentRows);

  if (enableMetamagicAutomationCheckbox && !enableMetamagicAutomationCheckbox.dataset.nasListenerAttached) {
      enableMetamagicAutomationCheckbox.dataset.nasListenerAttached = "true";
      enableMetamagicAutomationCheckbox.addEventListener("change", function () {
        toggleBuffSettingsVisibility(this.checked, metamagicDependentRows);
      });
  }
  });
}
