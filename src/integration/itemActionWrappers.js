import { MODULE } from "../common/module.js";
import { applyMetamagicDamageTransforms } from "../features/automation/metamagic/metamagicDamageTransforms.js";
import { GRAPPLE_CMB_MARKER } from "../features/automation/conditions/grappled/grappled.js";

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

export function registerItemActionWrappers() {
  if (!game.modules.get("lib-wrapper")?.active) {
    ui.notifications.error(game.i18n.format("NAS.integration.itemActions.libWrapperRequired", {
      moduleName: MODULE.NAME
    }));
    return;
  }

  libWrapper.register(
    MODULE.ID,
    "pf1.components.ItemAction.prototype.rollDamage",
    async function (wrapped, options = {}) {
      const rolls = await wrapped.apply(this, [options]);
      const rollData = options?.data ?? {};
      const metamagicData = rollData?.nasMeta ?? {};
      const transforms = Array.isArray(metamagicData.rollTransforms) ? metamagicData.rollTransforms : [];
      const intensifySourceParts = Array.isArray(metamagicData.intensifySourceParts)
        ? metamagicData.intensifySourceParts
        : [];
      if (!transforms.length) return rolls;

      for (let i = 0; i < rolls.length; i += 1) {
        const roll = rolls[i];
        const originalFormula = roll?.formula ?? "";
        if (!originalFormula) continue;
        const maximizeAlreadyApplied = Boolean(roll?.options?.maximize);
        const transformed = await applyMetamagicDamageTransforms({
          formula: originalFormula,
          transforms,
          rollData,
          maximizeAlreadyApplied,
          intensifySourceFormula: intensifySourceParts[i] ?? null
        });
        const nextFormula = transformed.formula;
        if (!nextFormula || nextFormula === originalFormula) continue;
        try {
          const nextRollOptions = {
            ...(roll.options ?? {}),
            nasLabelContext: {
              originalFormula,
              transformedFormula: nextFormula,
              transforms: Array.isArray(transformed?.applied) ? transformed.applied : [],
            },
          };
          const replaced = await new pf1.dice.DamageRoll(nextFormula, roll.data, nextRollOptions)
            .evaluate({ maximize: !!roll.options?.maximize, minimize: !!roll.options?.minimize });
          rolls[i] = replaced;
        } catch (_error) {}
      }

      return rolls;
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
