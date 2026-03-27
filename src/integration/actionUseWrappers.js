import { MODULE } from '../common/module.js';
import { handleBuffAutomation } from '../features/automation/buffs/buffs.js';
import { collectSpellActionData } from '../features/automation/utils/spellActionData.js';
import { applyMetamagicSelections } from '../features/automation/metamagic/applyMetamagic.js';
import { applyActionUseOverrides } from '../features/automation/utils/actionUseOverrides.js';

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
        await applyMetamagicSelections(this, shared.nasSpellContext);
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
}



