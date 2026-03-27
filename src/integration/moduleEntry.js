import { moduleConfig, registerNasSettings } from '../common/config.js';
import { MODULE } from '../common/module.js';
import { MigrationSettingsMenu, runSuiteMigrations } from '../common/migrations.js';
import { onRenderChatMessage } from '../features/buttons/healing.js';
import { addClusteredShotsButton } from '../features/buttons/clusteredshots.js';
import { handleCombatTrackerRender } from '../features/buttons/surpriseround.js';
import { DamageCommands } from '../features/commands/damageCommands.js';
import { handleCombatTurn, handleCombatRound, handleFlatFootedOnCombatStart } from '../features/automation/combat/combat.js';
import { applyChatRangeOverrides, registerChatRangeHoverOverrides } from '../features/automation/utils/chatRangeOverrides.js';
import { applyChatActivationOverrides } from '../features/automation/utils/chatActivationOverrides.js';
import { registerPersistentSpellSaveOverrides } from '../features/automation/utils/chatSaveOverrides.js';
import { applyEmpowerTooltipOverrides } from '../features/automation/metamagic/empowerSpell.js';
import { reorderTokenHUDConditions } from '../features/automation/utils/reorderTokenHUDConditions.js';
import { registerConditions, setupConditionsI18n } from '../features/automation/conditions/registry.js';
import { registerConditionFootnoteWrapper } from './conditionFootnoteWrappers.js';
import { registerActionUseWrapper } from './actionUseWrappers.js';
import { registerConfusionChatMessageHook } from '../features/automation/conditions/confusion/chatMessageHook.js';
import { handleConfusionOnCombatStart } from '../features/automation/conditions/confusion/confusion.js';
import { addGrappleCheckbox, addMetamagicCheckbox } from '../features/automation/utils/attackDialogControls.js';
import { handleNauseatedPreActionUse, handleNauseatedPreConcentration } from '../features/automation/conditions/nauseated/nauseated.js';
import { handleEntangledConcentration } from '../features/automation/conditions/entangled/entangled.js';
import { handleSqueezingPreActionUse } from '../features/automation/conditions/squeezing/squeezing.js';
import { checkNextTokenFlatFooted, restoreFlatFootedTracker, updateFlatFootedTracker } from '../features/automation/conditions/flatfooted/flatfooted.js';
import { handleDeadOnUpdate } from '../features/automation/conditions/dead/dead.js';
import { handleDisabledOnUpdate, handleDisabledStrenuousAction } from '../features/automation/conditions/disabled/disabled.js';
import { handleEnergyDrainOnUpdate } from '../features/automation/conditions/energydrain/energydrain.js';
import { handleGrappledActionUse, handleGrappledConcentration, handleGrappleResolution, isGrappleSelected } from '../features/automation/conditions/grappled/grappled.js';
import { handleUnconsciousOnUpdate } from '../features/automation/conditions/unconscious/unconscious.js';
import { registerLegacyDamageOverride } from '../features/automation/damage/legacyDamage.js';
import { registerSystemApplyDamage } from '../features/automation/damage/systemApplyDamage.js';
import { initializeSockets, initializeConditionIds } from './moduleSockets.js';
import { initItemAttackFlagCopy } from '../features/automation/damage/flagCopy.js';
import { handleConcealmentToggle } from '../features/automation/conditions/concealment/concealment.js';
import { registerDamageSettingsHandlebarsHelpers } from '../common/settings/damageSettingsForms.js';
import { registerDamageFootnoteHooks } from '../features/automation/damage/footnotes.js';

const LEGACY_MODULE_IDS = [MODULE.LEGACY_AD, MODULE.LEGACY_IC];

let _legacyChecked = false;
let _blockedByLegacyModules = false;
let _activeLegacy = [];

function ensureMigrationMenuRegistered() {
  try {
    const versionKey = `${MODULE.ID}.migrationVersion`;
    if (!game.settings.settings.has(versionKey)) {
      game.settings.register(MODULE.ID, "migrationVersion", {
        name: "Migration Version",
        scope: "world",
        config: false,
        type: String,
        default: ""
      });
    }

    const menuKey = `${MODULE.ID}.migrationTool`;
    if (!game.settings.settings.has(menuKey)) {
      game.settings.registerMenu(MODULE.ID, "migrationTool", {
        name: "Run Legacy Migration",
        label: "Open Migration Tool",
        hint: "Copy legacy module data into this module's settings and flags.",
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

  initializeConditionIds();
  registerConditionFootnoteWrapper(isGrappleSelected);
  registerActionUseWrapper();
});

Hooks.on("renderChatMessage", (app, html, data) => {
  if (isBlocked()) return;
  onRenderChatMessage(html);
  addClusteredShotsButton(html);
  registerConfusionChatMessageHook();
  applyChatRangeOverrides(app, html);
  applyChatActivationOverrides(app, html);
  applyEmpowerTooltipOverrides(html);
});

Hooks.on("renderChatLog", (_app, html) => {
  if (isBlocked()) return;
  registerChatRangeHoverOverrides(html);
  registerPersistentSpellSaveOverrides(html);
});

Hooks.once("socketlib.ready", () => {
  if (isBlocked()) return;
  initializeSockets();
});

Hooks.once("ready", () => {
  runSuiteMigrations();
  if (isBlocked()) return;

  moduleConfig.handleReadyHook();
  const didRegisterSystemDamage = registerSystemApplyDamage();
  if (!didRegisterSystemDamage) {
    registerLegacyDamageOverride();
  }
  DamageCommands.initialize();
  initItemAttackFlagCopy();
});

Hooks.on('renderTokenHUD', (app, html, data) => {
  if (isBlocked()) return;
  reorderTokenHUDConditions(html, data);
});

Hooks.on("pf1ToggleActorCondition", async (actor, conditionId, value) => {
  if (isBlocked()) return;
  await handleConcealmentToggle(actor, conditionId, value);
});

Hooks.on('renderAttackDialog', (dialog, html) => {
  if (isBlocked()) return;
  addGrappleCheckbox(dialog, html);
  addMetamagicCheckbox(dialog, html);
});

Hooks.on('combatStart', async (combat) => {
  if (isBlocked()) return;
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

Hooks.on('updateCombat', (combat, update, options, userId) => {
  if (isBlocked()) return;
  if (((combat.previous?.round === combat.current?.round) || (combat.previous?.round === 0)) &&
    ((combat.previous?.turn === combat.current?.turn) || (combat.previous?.turn === null)) &&
    (combat.previous?.tokenId === combat.turns[0]?.tokenId || combat.previous?.tokenId === null)
  ) return;
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
      updateFlatFootedTracker(combat);
    } else if (!game.user.isGM) {
      checkNextTokenFlatFooted(combat, combatData);
    }
  }
});

Hooks.on('renderCombatTracker', (app, html, data) => {
  if (isBlocked()) return;
  handleCombatTrackerRender(app, html, data);
});

Hooks.on('updateActor', async (actorDocument, change, options, userId) => {
  if (isBlocked()) return;
  if (!actorDocument.isOwner) return;

  let ownerUsers = game.users.filter(u => actorDocument.testUserPermission(u, 'OWNER'));

  let nonGMOwners = ownerUsers.filter(u => !u.isGM);

  let preferredUserId;
  if (nonGMOwners.length > 0) {
    preferredUserId = nonGMOwners[0].id;
  } else {
    preferredUserId = ownerUsers[0].id;
  }

  if (game.user.id !== preferredUserId && !game.user.isGM) return;

  await handleDisabledOnUpdate(actorDocument, change);
  await handleEnergyDrainOnUpdate(actorDocument, change);
  await handleUnconsciousOnUpdate(actorDocument, change);
  await handleDeadOnUpdate(actorDocument, change);
});

Hooks.on("pf1PreActionUse", async (action) => {
  if (isBlocked()) return;
  handleGrappledActionUse(action);
  handleNauseatedPreActionUse(action);
  handleSqueezingPreActionUse(action);
});

Hooks.on('pf1PreActorRollConcentration', (actor, rollContext) => {
  if (isBlocked()) return;
  return handleNauseatedPreConcentration(rollContext);
});

Hooks.on("pf1PostActionUse", async (action) => {
  if (isBlocked()) return;
  await handleGrappleResolution(action);

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
