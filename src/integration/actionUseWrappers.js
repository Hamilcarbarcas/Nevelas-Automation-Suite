import { MODULE } from '../common/module.js';
import { handleBuffAutomation } from '../features/automation/buffs/buffs.js';
import { collectSpellActionData } from '../features/automation/utils/spellActionData.js';
import { applyMetamagicSelections } from '../features/automation/metamagic/applyMetamagic.js';
import { applyActionUseOverrides } from '../features/automation/utils/actionUseOverrides.js';
import {
  createGrappleCmbAttackEntry,
  GRAPPLE_CMB_MARKER,
  isGrappleCmbAttack,
  isGrappleSelected,
} from '../features/automation/conditions/grappled/grappled.js';

function getManeuverActionType(itemAction) {
  const actionType = itemAction?.actionType;
  if (!actionType || !pf1?.config?.itemActionTypes?.[actionType]) return "mcman";

  const isRangedLike =
    actionType === "rwak" ||
    actionType === "rsak" ||
    actionType === "twak" ||
    actionType === "rcman";
  return isRangedLike ? "rcman" : "mcman";
}

function resolveCombatManeuverSizeBonus(actor, fallbackValue = 0) {
  const size = actor?.system?.traits?.actualSize ?? actor?.system?.traits?.size;
  if (size === undefined || size === null) return fallbackValue;

  const direct = pf1.config.sizeSpecialMods?.[size];
  if (Number.isFinite(direct)) return direct;

  const numericSize = Number(size);
  if (Number.isFinite(numericSize)) {
    const fromValues = Object.values(pf1.config.sizeSpecialMods ?? {})[numericSize];
    if (Number.isFinite(fromValues)) return fromValues;
  }
  return fallbackValue;
}

export function registerActionUseWrapper() {
  if (!game.modules.get("lib-wrapper")?.active) {
    ui.notifications.error(`${MODULE.NAME} requires the 'libWrapper' module. Please install and activate it.`);
    return;
  }

  libWrapper.register(
    MODULE.ID,
    "pf1.actionUse.ActionUse.prototype.process",
    async function(wrapped, ...args) {
      const itemType = this.item?.type;
      const itemSubType = this.item?.subType;
      const useCustomLogic = (itemType === "spell" || itemType === "consumable" || (itemType === "feat" && itemSubType === "classFeat"));
      if (useCustomLogic && game.settings.get(MODULE.ID, 'automaticBuffs')) {
        const shared = this.shared;
        let reqErr = await this.checkRequirements();
        if (reqErr > 0) return { err: pf1.actionUse.ERR_REQUIREMENT, code: reqErr };
        await this.autoSelectAmmo();
        this.getRollData();
        Hooks.callAll("pf1CreateActionUse", this);
        shared.fullAttack = true;
        await this.generateAttacks(true);
        let formData;
        const options = args[0] || {};
        if (options.skipDialog) {
          formData = {};
        } else {
          const dialog = new pf1.applications.AttackDialog(this);
          formData = await dialog.show();
          if (!formData) return;
        }
        this.formData = formData;
        this.shared.formData = formData;
        await this.alterRollData(formData);
        if (shared.action.ammo.type && shared.action.ammo?.cost > 0) {
          shared.attacks = shared.attacks.filter((o) => o.hasAmmo);
          if (shared.attacks.length === 0) {
            ui.notifications.error(game.i18n.localize("PF1.AmmoDepleted"));
            return { err: pf1.actionUse.ERR_REQUIREMENT, code: pf1.actionUse.ERR_REQUIREMENT.INSUFFICIENT_AMMO };
          }
        }
        if (!shared.fullAttack) shared.attacks = shared.attacks.slice(0, 1);
        await this.handleConditionals();
        await this.prepareChargeCost();
        if (shared.rollData.chargeCost != 0 && this.shared.action.uses?.perAttack) {
          const cost = shared.rollData.chargeCost;
          const charges = shared.item.charges;
          shared.attacks.forEach((atk, index) => {
            if (charges >= (index + 1) * cost) atk.chargeCost = cost;
            else atk.chargeCost = null;
          });
          shared.attacks = shared.attacks.filter((o) => o.chargeCost !== null);
          if (shared.attacks.length === 0) {
            ui.notifications.error(game.i18n.localize("PF1.ChargesDepleted"));
            return { err: pf1.actionUse.ERR_REQUIREMENT, code: pf1.actionUse.ERR_REQUIREMENT.INSUFFICIENT_CHARGES };
          }
        }
        reqErr = await this.checkAttackRequirements();
        if (reqErr > 0) return { err: pf1.actionUse.ERR_REQUIREMENT, code: reqErr };
        let measureResult;
        if (shared.useMeasureTemplate && canvas.scene) {
          measureResult = await this.promptMeasureTemplate();
          if (measureResult === null) return;
        }
        shared.nasSpellContext = await collectSpellActionData(this);
        const metamagicEnabled = game.settings.get(MODULE.ID, "enableMetamagicAutomation");
        if (metamagicEnabled) {
          await applyMetamagicSelections(this, shared.nasSpellContext);
        }
        const restoreOverrides = applyActionUseOverrides(this, shared.nasSpellContext);
        try {
          await this.getTargets();
          await this.generateChatAttacks();
          await this.addEffectNotes();
          await this.addFootnotes();
          if (Hooks.call("pf1PreActionUse", this) === false) {
            await measureResult?.delete();
            return;
          }
          await handleBuffAutomation(this);
          await this.executeScriptCalls();
          if (shared.scriptData?.reject) {
            await measureResult?.delete();
            return;
          }
          const premessage_promises = [];
          premessage_promises.push(this.handleDiceSoNice());
          const ammoCost = this.action.ammo.cost;
          if (ammoCost != 0) premessage_promises.push(this.subtractAmmo(ammoCost));
          let totalCost = shared.rollData?.chargeCost;
          if (this.action.uses.perAttack) {
            totalCost = this.shared.attacks.reduce((total, atk) => total + atk.chargeCost, 0);
          }
          if (totalCost != 0) {
            shared.totalChargeCost = totalCost;
            premessage_promises.push(this.item.addCharges(-totalCost));
          }
          if (shared.action.isSelfCharged)
            premessage_promises.push(shared.action.update({ "uses.self.value": shared.action.uses.self.value - 1 }));
          await Promise.all(premessage_promises);
          this.updateAmmoUsage();
          let result = Promise.resolve(null);
          await this.getMessageData();
          const rangeUnits = this.action?.range?.units;
          const increments = this.action?.range?.maxIncrements;
          const rollData = this.shared.rollData;
          const baseRange = this.action?.getRange?.({ type: "single", rollData });
          const minRange = this.action?.getRange?.({ type: "min", rollData });
          const maxRange = this.action?.getRange?.({ type: "max", rollData });
          const rangeOverride = {
            base: baseRange != null ? pf1.utils.convertDistanceBack(baseRange)[0] : null,
            min: minRange != null ? pf1.utils.convertDistanceBack(minRange)[0] : null,
            max: maxRange != null ? pf1.utils.convertDistanceBack(maxRange)[0] : null,
            units: rangeUnits ?? "",
            increments: increments ?? 1,
          };
          this.shared.chatData.flags ??= {};
          this.shared.chatData.flags[MODULE.ID] ??= {};
          this.shared.chatData.flags[MODULE.ID].actionOverrides ??= {};
          if (metamagicEnabled) {
            this.shared.chatData.flags[MODULE.ID].metamagic ??= {};
            if (shared.nasSpellContext?.metamagic?.persistent) {
              this.shared.chatData.flags[MODULE.ID].metamagic.persistent = true;
            }
            if (shared.nasSpellContext?.metamagic?.dazing) {
              this.shared.chatData.flags[MODULE.ID].metamagic.dazing = true;
              this.shared.chatData.flags[MODULE.ID].metamagic.dazingRounds =
                shared.nasSpellContext?.metamagic?.dazingRounds ?? 1;
              this.shared.chatData.flags[MODULE.ID].metamagic.dazingSpellName =
                shared.nasSpellContext?.metamagic?.dazingSpellName ?? "";
            }
          }
          if (Array.isArray(shared.targets) && shared.targets.length) {
            this.shared.chatData.flags[MODULE.ID].targets = shared.targets
              .map((target) => target?.document?.uuid ?? target?.uuid ?? target?.id)
              .filter(Boolean);
          }
          this.shared.chatData.flags[MODULE.ID].actionOverrides.range = rangeOverride;
          const activationOverride = shared.nasSpellContext?.activation;
          const extraFullRound = shared.nasSpellContext?.activationExtraFullRound;
          if (activationOverride || extraFullRound) {
            this.shared.chatData.flags[MODULE.ID].actionOverrides.activation = {
              activation: activationOverride,
              extraFullRound: Boolean(extraFullRound),
            };
          }
          if (shared.scriptData?.hideChat !== true) {
            result = this.postMessage();
          }
          await result;
        } finally {
          restoreOverrides();
        }
        if (game.settings.get("pf1", "clearTargetsAfterAttack") && game.user.targets.size) {
          game.user.updateTokenTargets([]);
          game.user.broadcastActivity({ targets: [] });
        }
        await this.executeScriptCalls("postUse");
        Hooks.callAll("pf1PostActionUse", this, this.shared.message ?? null);
        return this;
      } else {
        return wrapped.apply(this, args);
      }
    },
    "MIXED"
  );

  libWrapper.register(
    MODULE.ID,
    "pf1.actionUse.ActionUse.prototype.addAttacks",
    async function (wrapped, ...args) {
      await wrapped.apply(this, args);

      if (!isGrappleSelected(this) || !this?.action?.hasAttack) return;

      const existingGrappleAttack = (this.shared?.attacks ?? []).find((attack) => isGrappleCmbAttack(attack));
      if (existingGrappleAttack?.chatAttack?.attack) return;

      const ChatAttackClass = this.shared?.chatAttacks?.[0]?.constructor;
      if (!ChatAttackClass) return;

      const rollData = this.shared?.rollData;
      if (!rollData) return;

      const syntheticAttack = existingGrappleAttack ?? {
        ...createGrappleCmbAttackEntry(),
        abstract: true,
        ammo: null,
        chargeCost: null,
        chatAttack: null,
      };
      if (!existingGrappleAttack) this.shared.attacks.push(syntheticAttack);

      const attackIndex = this.shared.attacks.indexOf(syntheticAttack);
      rollData.attackCount = attackIndex + (this.shared?.skipAttacks ?? 0);

      const chatAttack = new ChatAttackClass(this.action, {
        label: syntheticAttack.label,
        rollData,
        targets: game.user.targets,
        actionUse: this,
      });

      const conditionalParts = this._getConditionalParts(syntheticAttack, { index: attackIndex });
      await chatAttack.addAttack({
        extraParts: [...(this.shared?.attackBonus ?? []), syntheticAttack.attackBonus],
        conditionalParts,
      });

      syntheticAttack.chatAttack = chatAttack;
      this.shared.chatAttacks.push(chatAttack);
      delete rollData.attackCount;
    },
    "MIXED"
  );

  libWrapper.register(
    MODULE.ID,
    "pf1.components.ItemAction.prototype.rollAttack",
    async function (wrapped, options = {}) {
      const extraParts = Array.isArray(options?.extraParts) ? options.extraParts : [];
      const hasMarker = extraParts.includes(GRAPPLE_CMB_MARKER);
      if (!hasMarker) return wrapped.apply(this, [options]);

      const filteredExtraParts = extraParts.filter((part) => part !== GRAPPLE_CMB_MARKER);
      const nextOptions = { ...options, extraParts: filteredExtraParts };
      const rollData = foundry.utils.deepClone(nextOptions.data ?? this.getRollData());
      nextOptions.data = rollData;

      const originalActionType = this.actionType;
      const originalManeuverType = this.maneuverType;
      const originalAttackBonus = this.attackBonus;
      const originalAbilityAttack = this.ability?.attack;
      const originalMasterwork = this.item?.system?.masterwork;
      let didOverrideMasterwork = false;
      const maneuverActionType = getManeuverActionType(this);
      const cmbAbility = this.actor?.system?.attributes?.cmbAbility ?? originalAbilityAttack;

      try {
        this.actionType = maneuverActionType;
        this.maneuverType = "grapple";
        this.attackBonus = "0";
        if (this.ability) this.ability.attack = cmbAbility;
        if (this.item?.system && this.item.system.masterwork !== undefined) {
          try {
            this.item.system.masterwork = false;
            didOverrideMasterwork = true;
          } catch (_err) {
          }
        }

        rollData.action ??= {};
        rollData.action.actionType = maneuverActionType;
        rollData.action.maneuverType = "grapple";
        rollData.action.attackBonus = "0";
        rollData.action.ability ??= {};
        rollData.action.ability.attack = cmbAbility;
        rollData.sizeBonus = resolveCombatManeuverSizeBonus(this.actor, rollData.sizeBonus);

        return await wrapped.apply(this, [nextOptions]);
      } finally {
        this.actionType = originalActionType;
        this.maneuverType = originalManeuverType;
        this.attackBonus = originalAttackBonus;
        if (this.ability) this.ability.attack = originalAbilityAttack;
        if (didOverrideMasterwork && this.item?.system && this.item.system.masterwork !== undefined) {
          try {
            this.item.system.masterwork = originalMasterwork;
          } catch (_err) {
          }
        }
      }
    },
    "MIXED"
  );
}



