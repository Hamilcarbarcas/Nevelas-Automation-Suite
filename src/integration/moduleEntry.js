import { moduleConfig, registerNasSettings } from '../common/config.js';
import { htmlElementFromRenderArg, onRenderChatMessageCompat } from '../common/foundryCompat.js';
import { MODULE } from '../common/module.js';
import { MigrationSettingsMenu, runSuiteMigrations } from '../common/migrations.js';
import { onRenderChatMessage } from '../features/buttons/healing.js';
import { addPerAttackHealthModeChip } from '../features/buttons/healthModeChip.js';
import { addClusteredShotsButton } from '../features/buttons/clusteredshots.js';
import { handleCombatTrackerRender } from '../features/buttons/surpriseround.js';
import { DamageCommands } from '../features/commands/damageCommands.js';
import { handleCombatTurn, handleCombatRound, handleFlatFootedOnCombatStart, skipIneligibleSurpriseCombatants } from '../features/automation/combat/combat.js';
import { applyChatRangeOverrides, registerChatRangeHoverOverrides } from '../features/automation/utils/chatRangeOverrides.js';
import { applyChatActivationOverrides } from '../features/automation/utils/chatActivationOverrides.js';
import { registerPersistentSpellSaveOverrides } from '../features/automation/utils/chatSaveOverrides.js';
import { applyEmpowerTooltipOverrides } from '../features/automation/metamagic/empowerSpell.js';
import { registerDiceTooltipEnhancer } from '../features/automation/utils/diceTooltipEnhancer.js';
import { registerHealthDeltaTextEnhancer } from '../features/automation/utils/healthDeltaText.js';
import { registerSpellPerfection } from '../features/automation/metamagic/feats/spellPerfection.js';
import { reorderTokenHUDConditions } from '../features/automation/utils/reorderTokenHUDConditions.js';
import { registerConditions, setupConditionsI18n } from '../features/automation/conditions/registry.js';
import { registerConditionFootnoteWrapper } from './conditionFootnoteWrappers.js';
import { registerActionUseWrapper } from './actionUseWrappers.js';
import { registerItemActionWrappers } from './itemActionWrappers.js';
import { registerConfusionChatMessageHook } from '../features/automation/conditions/confusion/chatMessageHook.js';
import { handleConfusionOnCombatStart } from '../features/automation/conditions/confusion/confusion.js';
import { addGrappleCheckbox, addMetamagicCheckbox } from '../features/automation/utils/attackDialogControls.js';
import { handleNauseatedPreActionUse, handleNauseatedPreConcentration } from '../features/automation/conditions/nauseated/nauseated.js';
import { handleEntangledConcentration } from '../features/automation/conditions/entangled/entangled.js';
import { handleSqueezingPreTokenUpdate, registerSqueezingTokenConfigFields } from '../features/automation/conditions/squeezing/squeezing.js';
import { handleSqueezedPreActionUse } from '../features/automation/conditions/squeezed/squeezed.js';
import { checkNextTokenFlatFooted, restoreFlatFootedTracker, updateFlatFootedTracker } from '../features/automation/conditions/flatfooted/flatfooted.js';
import { handleDeadOnUpdate } from '../features/automation/conditions/dead/dead.js';
import { handleDyingCombatUpdate, handleDyingOnUpdate, handleDyingPreActionUse } from '../features/automation/conditions/dying/dying.js';
import { handleDisabledOnUpdate, handleDisabledStrenuousAction } from '../features/automation/conditions/disabled/disabled.js';
import { handleEnergyDrainOnUpdate } from '../features/automation/conditions/energydrain/energydrain.js';
import { handleGrappledActionUse, handleGrappledConcentration, handleGrappleResolution, isGrappleSelected } from '../features/automation/conditions/grappled/grappled.js';
import { handleUnconsciousOnUpdate } from '../features/automation/conditions/unconscious/unconscious.js';
import { handleWoundsVigorActionTax } from '../features/automation/conditions/woundsvigor/actionTax.js';
import { handleWoundsVigorThresholdSync } from '../features/automation/conditions/woundsvigor/thresholdSync.js';
import { registerSystemApplyDamage } from '../features/automation/damage/systemApplyDamage.js';
import { initializeSockets, initializeConditionIds } from './moduleSockets.js';
import { initItemAttackFlagCopy } from '../features/automation/damage/flagCopy.js';
import { registerConcealedConditionWrappers } from '../features/automation/conditions/concealed/concealed.js';
import { registerDamageSettingsHandlebarsHelpers } from '../common/settings/damageSettingsForms.js';
import { registerDamageFootnoteHooks } from '../features/automation/damage/footnotes.js';
import { registerReactiveItemSheet } from '../features/automation/damage/reactiveItemSheet.js';
import {
  applyMirrorImageChatContent,
  isMirrorImageBuff,
  refreshMirrorImageSceneTokenEffects,
  refreshMirrorImageTokenEffects,
  registerMirrorImageTokenEffectBadges,
  renderMirrorImageChatControls
} from '../features/automation/buffs/mirrorImage.js';
import { hasHpUpdate } from '../features/automation/utils/healthUpdates.js';
import { handleHtkCombatUpdate } from '../features/automation/utils/hardToKillCombat.js';
import {
  applySpellAbilityMinimumSheetVisibility,
  enforceSpellAbilityMinimumOnActionUse,
  enforceSpellAbilityMinimumOnPreCreateItem,
  enforceSpellAbilityMinimumOnPreUpdateItem
} from '../features/automation/spellcasting/abilityLimit.js';
import {
  registerEldritchResearcherItemHook,
  registerSpontaneousMetafocusItemHook,
  registerSpellPerfectionItemHook
} from '../features/automation/metamagic/feats/index.js';
import {
  registerMagicalLineageItemHook,
  registerTransmuterOfKoradaItemHook,
  registerWayangSpellhunterItemHook
} from '../features/automation/metamagic/traits/index.js';

const LEGACY_MODULE_IDS = [MODULE.LEGACY_AD, MODULE.LEGACY_IC];

let _legacyChecked = false;
let _blockedByLegacyModules = false;
let _activeLegacy = [];

function ensureMigrationMenuRegistered() {
  try {
    const versionKey = `${MODULE.ID}.migrationVersion`;
    if (!game.settings.settings.has(versionKey)) {
      game.settings.register(MODULE.ID, "migrationVersion", {
        name: game.i18n.localize("NAS.settings.migrationVersion.name"),
        scope: "world",
        config: false,
        type: String,
        default: ""
      });
    }

    const menuKey = `${MODULE.ID}.migrationTool`;
    if (!game.settings.settings.has(menuKey)) {
      game.settings.registerMenu(MODULE.ID, "migrationTool", {
        name: game.i18n.localize("NAS.settings.migrationTool.name"),
        label: game.i18n.localize("NAS.settings.migrationTool.label"),
        hint: game.i18n.localize("NAS.settings.migrationTool.hint"),
        icon: "fas fa-database",
        type: MigrationSettingsMenu,
        restricted: true
      });
    }
  } catch (err) {
    console.error(`[${MODULE.ID}] Failed to register migration tools`, err);
  }
}

function ensureMigrationPersistenceSettingsRegistered() {
  try {
    if (!game.settings.settings.has(`${MODULE.ID}.customDamageTypes`)) {
      game.settings.register(MODULE.ID, "customDamageTypes", {
        scope: "world",
        config: false,
        type: Array,
        default: [],
        requiresReload: true
      });
    }

    if (!game.settings.settings.has(`${MODULE.ID}.translations`)) {
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
    }

    if (!game.settings.settings.has(`${MODULE.ID}.damageTypePriority`)) {
      game.settings.register(MODULE.ID, "damageTypePriority", {
        name: game.i18n?.localize?.("NAS.settings.damageTypePriority.name") ?? "Custom Damage Type Priority",
        hint: game.i18n?.localize?.("NAS.settings.damageTypePriority.hint") ?? "",
        default: JSON.stringify([[], ["magic"], [], ["alchemicalSilver", "coldIron", "mithral", "nexavarianSteel", "sunsilver"], ["adamantine"], ["lawful", "chaotic", "good", "evil"], ["epic"]]),
        scope: "world",
        type: String,
        config: false
      });
    }
  } catch (err) {
    console.error(`[${MODULE.ID}] Failed to register migration persistence settings`, err);
  }
}

function computeLegacyBlockIfPossible() {
  if (_legacyChecked) return;
  const mods = globalThis.game?.modules;
  if (!mods || typeof mods.get !== "function") return;

  _legacyChecked = true;
  _activeLegacy = LEGACY_MODULE_IDS
    .map((id) => mods.get(id))
    .filter((m) => m?.active)
    .map((m) => ({
      id: m.id,
      title: m.title ?? m.id
    }));

  if (_activeLegacy.length > 0) {
    _blockedByLegacyModules = true;
    console.error(`[${MODULE.ID}] Legacy predecessor modules detected; NAS will be disabled.`, { activeLegacy: _activeLegacy });
    Hooks.once('ready', () => {
      try {
        const legacyList = _activeLegacy.map((m) => m.title).join(', ');
        const key = 'NAS.legacy.blocked';
        const fallback =
          `Nevela's Automation Suite is incompatible with the legacy module(s) ${legacyList}. ` +
          `These modules are no longer supported. Disable them to use Nevela's Automation Suite. ` +
          `(The legacy migration tool remains available.)`;
        const i18n = globalThis.game?.i18n;
        const msg = (i18n?.format && (!i18n.has || i18n.has(key))) ? i18n.format(key, { legacyList }) : fallback;
        globalThis.ui?.notifications?.error?.(msg && msg !== key ? msg : fallback, { permanent: true });
      } catch {
      }
    });
  }
}

function isBlocked() {
  computeLegacyBlockIfPossible();
  return _blockedByLegacyModules;
}

function isConditionAutomationEnabled() {
  const key = `${MODULE.ID}.automateConditions`;
  if (!globalThis.game?.settings?.settings?.has?.(key)) return true;
  return game.settings.get(MODULE.ID, "automateConditions");
}

function handleNasRenderChatMessage(app, html, data) {
  if (isBlocked()) return;
  const root = htmlElementFromRenderArg(html);
  if (!root) return;
  addPerAttackHealthModeChip(root);
  onRenderChatMessage(root);
  addClusteredShotsButton(root);
  applyChatRangeOverrides(app, root);
  applyChatActivationOverrides(app, root);
  renderMirrorImageChatControls(app, root);
  if (game.settings.get(MODULE.ID, "enableMetamagicAutomation")) {
    applyEmpowerTooltipOverrides(root);
  }
}

Hooks.on('pf1RegisterDamageTypes', (registry) => {
  ensureMigrationMenuRegistered();
  computeLegacyBlockIfPossible();
  if (isBlocked()) {
    ensureMigrationPersistenceSettingsRegistered();
    return;
  }
  moduleConfig.handleRegistryHook(registry);
});

Hooks.once("setup", (...args) => {
  if (isBlocked()) return;
  moduleConfig.handleSetupHook(...args);
});

Hooks.on("pf1RegisterConditions", (registry) => {
  if (isBlocked()) return;
  registerConditions(registry);
});

Hooks.on('little-helper.i18n', (t) => {
  if (isBlocked()) return;
  setupConditionsI18n(t);
});

Hooks.once("init", () => {
  ensureMigrationMenuRegistered();
  computeLegacyBlockIfPossible();
  if (isBlocked()) {
    ensureMigrationPersistenceSettingsRegistered();
    return;
  }

  registerNasSettings();

  registerDamageSettingsHandlebarsHelpers();
  registerDamageFootnoteHooks();
  registerReactiveItemSheet();
  registerSqueezingTokenConfigFields();

  initializeConditionIds();
  registerConditionFootnoteWrapper(isGrappleSelected);
  registerActionUseWrapper();
  registerItemActionWrappers();
  registerConcealedConditionWrappers();
  registerConfusionChatMessageHook();
  registerMirrorImageTokenEffectBadges();
  registerDiceTooltipEnhancer();
  registerHealthDeltaTextEnhancer();
  registerSpellPerfection();
  registerEldritchResearcherItemHook();
  registerSpontaneousMetafocusItemHook();
  registerSpellPerfectionItemHook();
  registerMagicalLineageItemHook();
  registerTransmuterOfKoradaItemHook();
  registerWayangSpellhunterItemHook();
});

onRenderChatMessageCompat(handleNasRenderChatMessage);

Hooks.on("renderChatLog", (_app, html) => {
  if (isBlocked()) return;
  registerChatRangeHoverOverrides(html);
  registerPersistentSpellSaveOverrides(html);
});

Hooks.on("preCreateChatMessage", (message) => {
  if (isBlocked()) return;
  applyMirrorImageChatContent(message);
});

Hooks.once("socketlib.ready", () => {
  if (isBlocked()) return;
  initializeSockets();
});

Hooks.once("ready", () => {
  runSuiteMigrations();
  if (isBlocked()) return;

  moduleConfig.handleReadyHook();
  registerSystemApplyDamage();
  DamageCommands.initialize();
  initItemAttackFlagCopy();
});

Hooks.on("canvasReady", () => {
  if (isBlocked()) return;
  refreshMirrorImageSceneTokenEffects();
});

Hooks.on("createToken", (tokenDocument) => {
  if (isBlocked()) return;
  refreshMirrorImageTokenEffects(tokenDocument?.actor);
});

Hooks.on("updateToken", (tokenDocument) => {
  if (isBlocked()) return;
  refreshMirrorImageTokenEffects(tokenDocument?.actor);
});

Hooks.on("createItem", (item) => {
  if (isBlocked() || !isMirrorImageBuff(item)) return;
  refreshMirrorImageTokenEffects(item.actor);
});

Hooks.on("updateItem", (item) => {
  if (isBlocked() || !isMirrorImageBuff(item)) return;
  refreshMirrorImageTokenEffects(item.actor);
});

Hooks.on("deleteItem", (item) => {
  if (isBlocked() || !isMirrorImageBuff(item)) return;
  refreshMirrorImageTokenEffects(item.actor);
});

Hooks.on('renderTokenHUD', (app, html, data) => {
  if (isBlocked()) return;
  reorderTokenHUDConditions(html, data);
});

Hooks.on('renderAttackDialog', (dialog, html) => {
  if (isBlocked()) return;
  addGrappleCheckbox(dialog, html);
  if (game.settings.get(MODULE.ID, "enableMetamagicAutomation")) {
    addMetamagicCheckbox(dialog, html);
  }
});

Hooks.on("renderActorSheet", (app, html) => {
  if (isBlocked()) return;
  applySpellAbilityMinimumSheetVisibility(app, html);
});

Hooks.on('preUpdateToken', (tokenDocument, updateData, options, userId) => {
  if (isBlocked()) return;
  if (!isConditionAutomationEnabled()) return;
  return handleSqueezingPreTokenUpdate(tokenDocument, updateData, options, userId);
});

Hooks.on('combatStart', async (combat) => {
  if (isBlocked()) return;
  if (!isConditionAutomationEnabled()) return;
  restoreFlatFootedTracker(combat);

  const turnOrder = combat.turns;
  const isSurprise = combat.getFlag(MODULE.ID, 'isSurprise') || false;
  const highestInitiative = Math.max(...combat.combatants.map(c => c.initiative));

  for (const combatant of combat.combatants) {
    const token = canvas.tokens.get(combatant.tokenId);
    if (!token) continue;

    const turnIndex = turnOrder.findIndex(turn => turn.tokenId === combatant.tokenId);

    if (turnIndex !== -1) {
      if (game.settings.get(MODULE.ID, 'autoApplyFF')) {
        await handleFlatFootedOnCombatStart(combat, combatant, token, turnIndex, highestInitiative, isSurprise);
      }

      if (game.settings.get(MODULE.ID, 'handleConfused')) {
        await handleConfusionOnCombatStart(combatant, token, turnOrder);
      }
    }
  }

  updateFlatFootedTracker(combat);
});

Hooks.on('updateCombat', async (combat, update, options, userId) => {
  if (isBlocked()) return;
  const conditionAutomationEnabled = isConditionAutomationEnabled();
  const startupShortCircuit = (((combat.previous?.round === combat.current?.round) || (combat.previous?.round === 0)) &&
    ((combat.previous?.turn === combat.current?.turn) || (combat.previous?.turn === null)) &&
    (combat.previous?.tokenId === combat.turns[0]?.tokenId || combat.previous?.tokenId === null)
  );

  const hasTurnUpdate = update?.turn !== undefined && update?.turn !== null;
  const hasRoundUpdate = update?.round !== undefined && update?.round !== null;
  const isTurnOrRoundUpdate = hasTurnUpdate || hasRoundUpdate;

  if (conditionAutomationEnabled && game.user.isGM && userId === game.user.id && isTurnOrRoundUpdate && options?.nasSurpriseSkip !== true) {
    const didSkip = await skipIneligibleSurpriseCombatants(combat);
    if (didSkip) return;
  }

  if (startupShortCircuit) return;

  if (update.round !== undefined && game.user.isGM && userId === game.user.id) {
    handleCombatRound(combat, update.round);
  }

  if ((update.turn !== undefined && update.turn !== null && combat.combatant) ||
    (update.round !== undefined && combat.turns.length === 1 && combat.combatant)
  ) {
    const combatData = {
      combatantId: combat.combatant.id,
      tokenId: combat.combatant.token.id,
      turn: combat.turn,
      round: combat.round
    };

    if (game.user.isGM) {
      handleCombatTurn(combat, combatData);
      if (conditionAutomationEnabled) {
        updateFlatFootedTracker(combat);
        handleHtkCombatUpdate(combat, update);
        await handleDyingCombatUpdate(combat, update);
      }
    } else if (!game.user.isGM) {
      if (conditionAutomationEnabled) {
        checkNextTokenFlatFooted(combat, combatData);
      }
    }
  }
});

Hooks.on('renderCombatTracker', (app, html, data) => {
  if (isBlocked()) return;
  handleCombatTrackerRender(app, html, data);
});

Hooks.on('preUpdateActor', (actorDocument, change, options) => {
  try {
    if (hasHpUpdate(change)) {
      options._nasPrevHp = actorDocument.system?.attributes?.hp?.value;
    }
  } catch (_err) {
    // ignore
  }
});

Hooks.on('updateActor', async (actorDocument, change, options, userId) => {
  if (isBlocked()) return;
  if (!isConditionAutomationEnabled()) return;
  if (!actorDocument.isOwner) return;

  const ownerUsers = game.users.filter((u) => actorDocument.testUserPermission(u, 'OWNER'));
  const activeNonGMOwner = ownerUsers.find((u) => !u.isGM && u.active);
  const activeOwner = ownerUsers.find((u) => u.active);
  const preferredUserId = activeNonGMOwner?.id ?? activeOwner?.id ?? ownerUsers[0]?.id ?? null;
  if (!preferredUserId || game.user.id !== preferredUserId) return;

  await handleDisabledOnUpdate(actorDocument, change);
  await handleEnergyDrainOnUpdate(actorDocument, change);
  await handleUnconsciousOnUpdate(actorDocument, change, options);
  await handleDyingOnUpdate(actorDocument, change);
  await handleDeadOnUpdate(actorDocument, change);
  await handleWoundsVigorThresholdSync(actorDocument, change);
});

Hooks.on("pf1PreActionUse", async (action) => {
  if (isBlocked()) return;
  if (enforceSpellAbilityMinimumOnActionUse(action) === false) return;
  if (!isConditionAutomationEnabled()) return;
  handleGrappledActionUse(action);
  handleNauseatedPreActionUse(action);
  handleSqueezedPreActionUse(action);
  handleDyingPreActionUse(action);
});

Hooks.on("preCreateItem", (item, data, options, userId) => {
  if (isBlocked()) return;
  return enforceSpellAbilityMinimumOnPreCreateItem(item, data, userId);
});

Hooks.on("preUpdateItem", (item, change, options, userId) => {
  if (isBlocked()) return;
  return enforceSpellAbilityMinimumOnPreUpdateItem(item, change, userId);
});

Hooks.on('pf1PreActorRollConcentration', (actor, rollContext) => {
  if (isBlocked()) return;
  if (!isConditionAutomationEnabled()) return;
  return handleNauseatedPreConcentration(rollContext);
});

Hooks.on("pf1PostActionUse", async (action) => {
  if (isBlocked()) return;
  if (!isConditionAutomationEnabled()) return;
  await handleGrappleResolution(action);
  await handleWoundsVigorActionTax(action);

  if (game.settings.get(MODULE.ID, 'disableAtZeroHP') ||
    game.settings.get(MODULE.ID, 'entangledGrappledHandling') !== 'disabled'
  ) {
    const itemSource = action.item;

    if (game.settings.get(MODULE.ID, 'disableAtZeroHP')) {
      await handleDisabledStrenuousAction(action);
    }

    if (game.settings.get(MODULE.ID, 'entangledGrappledHandling') !== 'disabled') {
      if (itemSource.type == "spell") {
        await handleEntangledConcentration(action);
        await handleGrappledConcentration(action);
      }
    }
  }
});
