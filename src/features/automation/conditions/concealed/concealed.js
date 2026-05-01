import { MODULE } from '../../../../common/module.js';

export const CONCEALED_CONDITION_ID = "concealed";
export const CONCEALED_VARIANT_FLAG = "concealedVariant"; 
const INVISIBLE_CONDITION_ID = "invisible";

export function actorHasBlindFight(actor) {
  return actor?.items?.some(i => i.type === "feat" && i.name?.toLowerCase() === "blind-fight") ?? false;
}

export function getConcealedVariant(actor) {
  if (!actor) return null;
  for (const ae of actor.effects ?? []) {
    if (!ae?.statuses?.has?.(CONCEALED_CONDITION_ID)) continue;
    const v = ae.getFlag?.(MODULE.ID, CONCEALED_VARIANT_FLAG);
    return (v === "total" || v === "normal") ? v : "normal";
  }
  return null;
}

function getLocalizedConditionName() {
  const key = "NAS.conditions.list.concealed.label";
  const t = game?.i18n?.localize?.(key);
  return (t && t !== key) ? t : "Concealed";
}

function getVariantLabel(variant) {
  const key = variant === "total"
    ? "NAS.conditions.main.ConcealedTotal"
    : "NAS.conditions.main.ConcealedNormal";
  const t = game?.i18n?.localize?.(key);
  if (t && t !== key) return t;
  return variant === "total" ? "Total" : "Normal";
}

async function promptConcealedVariant() {
  return new Promise(resolve => {
    const buttons = {
      normal: {
        label: game.i18n.localize('NAS.conditions.main.ConcealedNormal'),
        callback: () => resolve("normal")
      },
      total: {
        label: game.i18n.localize('NAS.conditions.main.ConcealedTotal'),
        callback: () => resolve("total")
      },
      cancel: {
        label: game.i18n.localize('NAS.common.buttons.cancel') ?? "Cancel",
        callback: () => resolve(null)
      }
    };

    new Dialog({
      title: game.i18n.localize('NAS.conditions.main.ConcealedDialogTitle'),
      content: `<p>${game.i18n.localize('NAS.conditions.list.concealed.description')}</p>`,
      buttons,
      default: "normal"
    }).render(true);
  });
}

function clonePlainObject(value) {
  const fu = globalThis.foundry?.utils;
  if (fu?.getType?.(value) === "Object" && typeof fu.duplicate === "function") {
    return fu.duplicate(value);
  }
  if (value && typeof value === "object") {
    return { ...value };
  }
  return {};
}

export function registerConcealedConditionWrappers() {
  if (!game.modules.get("lib-wrapper")?.active) {
    console.warn(`${MODULE.ID} | libWrapper missing; concealed variant prompt disabled.`);
    return;
  }

  libWrapper.register(
    MODULE.ID,
    "pf1.documents.actor.ActorPF.prototype.setCondition",
    async function (wrapped, conditionId, enabled, context) {
      try {
        if (conditionId !== CONCEALED_CONDITION_ID && conditionId !== INVISIBLE_CONDITION_ID) {
          return await wrapped(conditionId, enabled, context);
        }

        if (enabled === false) {
          return await wrapped(conditionId, enabled, context);
        }

        const enabledIsObject = globalThis.foundry?.utils?.getType?.(enabled) === "Object";
        const incoming = enabledIsObject ? enabled : {};

        if (conditionId === INVISIBLE_CONDITION_ID) {
          const next = clonePlainObject(incoming);
          next.flags ??= {};
          next.flags[MODULE.ID] ??= {};
          next.flags[MODULE.ID][CONCEALED_VARIANT_FLAG] = "total";
          return await wrapped(conditionId, next, context);
        }

        const already = incoming?.flags?.[MODULE.ID]?.[CONCEALED_VARIANT_FLAG];
        if (already === "normal" || already === "total") {
          return await wrapped(conditionId, enabled, context);
        }

        const variant = await promptConcealedVariant();
        if (!variant) {
          const currentState = this?.statuses?.has?.(conditionId) ?? false;
          return await wrapped(conditionId, currentState, context);
        }

        const next = clonePlainObject(incoming);
        next.flags ??= {};
        next.flags[MODULE.ID] ??= {};
        next.flags[MODULE.ID][CONCEALED_VARIANT_FLAG] = variant;

        const baseName = getLocalizedConditionName();
        next.name = (variant === "total") ? `${baseName} (${getVariantLabel(variant)})` : baseName;

        return await wrapped(conditionId, next, context);
      } catch (err) {
        console.error(`${MODULE.ID} | Concealed condition wrapper failed`, err);
        return await wrapped(conditionId, enabled, context);
      }
    },
    "WRAPPER"
  );
}
