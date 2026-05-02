import { MODULE } from '../common/module.js';
import { htmlElementFromRenderArg, onRenderChatMessageCompat } from '../common/foundryCompat.js';
import { onRenderChatMessage } from '../features/buttons/healing.js';
import { addPerAttackHealthModeChip } from '../features/buttons/healthModeChip.js';
import { addClusteredShotsButton } from '../features/buttons/clusteredshots.js';
import { handleCombatTrackerRender } from '../features/buttons/surpriseround.js';
import { DamageCommands } from '../features/commands/damageCommands.js';
import { handleCombatTurn, handleCombatRound, handleFlatFootedOnCombatStart, skipIneligibleSurpriseCombatants } from '../features/automation/combat/combat.js';
import { registerConcealedConditionWrappers } from '../features/automation/conditions/concealed/concealed.js';
import { initializeSockets, initializeConditionIds } from './moduleSockets.js';
import { initItemAttackFlagCopy } from '../features/automation/damage/flagCopy.js';
import { registerConfusionChatMessageHook } from '../features/automation/conditions/confusion/chatMessageHook.js';
import { handleConfusionOnCombatStart } from '../features/automation/conditions/confusion/confusion.js';
import { addGrappleCheckbox, addMetamagicCheckbox } from '../features/automation/utils/attackDialogControls.js';
import { applyChatRangeOverrides, registerChatRangeHoverOverrides } from '../features/automation/utils/chatRangeOverrides.js';
import { applyChatActivationOverrides } from '../features/automation/utils/chatActivationOverrides.js';
import { registerPersistentSpellSaveOverrides } from '../features/automation/utils/chatSaveOverrides.js';
import { applyEmpowerTooltipOverrides } from '../features/automation/metamagic/empowerSpell.js';
import { registerDiceTooltipEnhancer } from '../features/automation/utils/diceTooltipEnhancer.js';
import { registerHealthDeltaTextEnhancer } from '../features/automation/utils/healthDeltaText.js';
import { registerSpellPerfectionCkl } from '../features/automation/metamagic/feats/spellPerfection.js';
import { reorderTokenHUDConditions } from '../features/automation/utils/reorderTokenHUDConditions.js';
import { registerConditions, setupConditionsI18n } from '../features/automation/conditions/registry.js';
import { registerConditionFootnoteWrapper } from './conditionFootnoteWrappers.js';
import { registerActionUseWrapper } from './actionUseWrappers.js';
import { registerItemActionWrappers } from './itemActionWrappers.js';
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
import { registerDamageSettingsHandlebarsHelpers } from '../common/settings/damageSettingsForms.js';
import { registerDamageFootnoteHooks } from '../features/automation/damage/footnotes.js';
import { hasHpUpdate } from '../features/automation/utils/healthUpdates.js';
import { handleHtkCombatUpdate } from '../features/automation/utils/hardToKillCombat.js';
import { registerTransmuterOfKoradaItemHook } from '../features/automation/metamagic/traits/index.js';

function isConditionAutomationEnabled() {
  const key = `${MODULE.ID}.automateConditions`;
  if (!globalThis.game?.settings?.settings?.has?.(key)) return true;
  return game.settings.get(MODULE.ID, "automateConditions");
}

export function registerNasModule() {
  Hooks.once("init", () => {
    registerDamageSettingsHandlebarsHelpers();
    registerDamageFootnoteHooks();
    registerSqueezingTokenConfigFields();
    initializeConditionIds();
    registerConditionFootnoteWrapper(isGrappleSelected);
    registerActionUseWrapper();
    registerItemActionWrappers();
    registerConcealedConditionWrappers();
    registerConfusionChatMessageHook();
    registerDiceTooltipEnhancer();
    registerHealthDeltaTextEnhancer();
    registerSpellPerfectionCkl();
    registerTransmuterOfKoradaItemHook();
  });

  onRenderChatMessageCompat((app, html, data) => {
    const root = htmlElementFromRenderArg(html);
    if (!root) return;
    addPerAttackHealthModeChip(root);
    onRenderChatMessage(root);
    addClusteredShotsButton(root);
    applyChatRangeOverrides(app, root);
    applyChatActivationOverrides(app, root);
    if (game.settings.get(MODULE.ID, "enableMetamagicAutomation")) {
      applyEmpowerTooltipOverrides(root);
    }
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

    registerSystemApplyDamage();
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

  Hooks.on('renderAttackDialog', (dialog, html) => {
    addGrappleCheckbox(dialog, html);
    if (game.settings.get(MODULE.ID, "enableMetamagicAutomation")) {
      addMetamagicCheckbox(dialog, html);
    }
  });

  Hooks.on('preUpdateToken', (tokenDocument, updateData, options, userId) => {
    if (!isConditionAutomationEnabled()) return;
    return handleSqueezingPreTokenUpdate(tokenDocument, updateData, options, userId);
  });

  Hooks.on('combatStart', async (combat) => {
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
    if (!isConditionAutomationEnabled()) return;
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
    await handleUnconsciousOnUpdate(actorDocument, change, options);
    await handleDyingOnUpdate(actorDocument, change);
    await handleDeadOnUpdate(actorDocument, change);
    await handleWoundsVigorThresholdSync(actorDocument, change);
  });

  Hooks.on("pf1PreActionUse", async (action) => {
    if (!isConditionAutomationEnabled()) return;
    handleGrappledActionUse(action);
    handleNauseatedPreActionUse(action);
    handleSqueezedPreActionUse(action);
    handleDyingPreActionUse(action);
  });

  Hooks.on('pf1PreActorRollConcentration', (actor, rollContext) => {
    if (!isConditionAutomationEnabled()) return;
    return handleNauseatedPreConcentration(rollContext);
  });

  Hooks.on("pf1PostActionUse", async (action) => {
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
}

