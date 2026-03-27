import { MODULE } from '../common/module.js';
import { onRenderChatMessage } from '../features/buttons/healing.js';
import { addClusteredShotsButton } from '../features/buttons/clusteredshots.js';
import { handleCombatTrackerRender } from '../features/buttons/surpriseround.js';
import { DamageCommands } from '../features/commands/damageCommands.js';
import { handleCombatTurn, handleCombatRound, handleFlatFootedOnCombatStart } from '../features/automation/combat/combat.js';
import { handleConcealmentToggle } from '../features/automation/conditions/concealment/concealment.js';
import { initializeSockets, initializeConditionIds } from './moduleSockets.js';
import { initItemAttackFlagCopy } from '../features/automation/damage/flagCopy.js';
import { registerConfusionChatMessageHook } from '../features/automation/conditions/confusion/chatMessageHook.js';
import { handleConfusionOnCombatStart } from '../features/automation/conditions/confusion/confusion.js';
import { addGrappleCheckbox, addMetamagicCheckbox } from '../features/automation/utils/attackDialogControls.js';
import { applyChatRangeOverrides, registerChatRangeHoverOverrides } from '../features/automation/utils/chatRangeOverrides.js';
import { applyChatActivationOverrides } from '../features/automation/utils/chatActivationOverrides.js';
import { registerPersistentSpellSaveOverrides } from '../features/automation/utils/chatSaveOverrides.js';
import { applyEmpowerTooltipOverrides } from '../features/automation/metamagic/empowerSpell.js';
import { reorderTokenHUDConditions } from '../features/automation/utils/reorderTokenHUDConditions.js';
import { registerConditions, setupConditionsI18n } from '../features/automation/conditions/registry.js';
import { registerConditionFootnoteWrapper } from './conditionFootnoteWrappers.js';
import { registerActionUseWrapper } from './actionUseWrappers.js';
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
import { registerDamageSettingsHandlebarsHelpers } from '../common/settings/damageSettingsForms.js';
import { registerDamageFootnoteHooks } from '../features/automation/damage/footnotes.js';

export function registerNasModule() {
  Hooks.once("init", () => {
    registerDamageSettingsHandlebarsHelpers();
    registerDamageFootnoteHooks();
    initializeConditionIds();
    registerConditionFootnoteWrapper(isGrappleSelected);
    registerActionUseWrapper();
  });

  Hooks.on("renderChatMessage", (app, html, data) => {
    onRenderChatMessage(html);
    addClusteredShotsButton(html);
    registerConfusionChatMessageHook();
    applyChatRangeOverrides(app, html);
    applyChatActivationOverrides(app, html);
    applyEmpowerTooltipOverrides(html);
  });

  Hooks.on("renderChatLog", (_app, html) => {
    registerChatRangeHoverOverrides(html);
    registerPersistentSpellSaveOverrides(html);
  });

  Hooks.once("socketlib.ready", () => {
    initializeSockets();
  });

  Hooks.once("ready", () => {
    import('../common/config.js')
      .then(({ moduleConfig }) => moduleConfig.handleReadyHook())
      .catch((err) => console.error(`[${MODULE.ID}] Ready hook failed`, err));

    const didRegisterSystemDamage = registerSystemApplyDamage();
    if (!didRegisterSystemDamage) {
      registerLegacyDamageOverride();
    }

    DamageCommands.initialize();
    initItemAttackFlagCopy();
  });

  Hooks.on("pf1RegisterConditions", (registry) => {
    registerConditions(registry);
  });

  Hooks.on('little-helper.i18n', (t) => {
    setupConditionsI18n(t);
  });

  Hooks.on('renderTokenHUD', (app, html, data) => {
    reorderTokenHUDConditions(html, data);
  });

  Hooks.on("pf1ToggleActorCondition", async (actor, conditionId, value) => {
    await handleConcealmentToggle(actor, conditionId, value);
  });

  Hooks.on('renderAttackDialog', (dialog, html) => {
    addGrappleCheckbox(dialog, html);
    addMetamagicCheckbox(dialog, html);
  });

  Hooks.on('combatStart', async (combat) => {
    console.log(combat);
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
    handleCombatTrackerRender(app, html, data);
  });

  Hooks.on('updateActor', async (actorDocument, change, options, userId) => {
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
    handleGrappledActionUse(action);
    handleNauseatedPreActionUse(action);
    handleSqueezingPreActionUse(action);
  });

  Hooks.on('pf1PreActorRollConcentration', (actor, rollContext) => {
    return handleNauseatedPreConcentration(rollContext);
  });

  Hooks.on("pf1PostActionUse", async (action) => {
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
}

