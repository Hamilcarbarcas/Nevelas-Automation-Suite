import { MODULE } from "../../../../common/module.js";
import { getPrimaryHealthValue, hasPrimaryHealthUpdate } from "../../utils/healthUpdates.js";
import { isWoundsVigorActive } from "../../utils/woundsVigor.js";

function isDyingStateRelevant(actor) {
  if (!actor) return false;
  if (actor.statuses?.has?.("dead")) return false;
  if (isWoundsVigorActive(actor)) return false;
  const hp = Number(actor.system?.attributes?.hp?.value ?? 0) || 0;
  return hp < 0;
}

function getRollDataFromResult(result) {
  const first = result?.rolls?.[0];
  if (!first) return null;
  if (typeof first === "string") {
    try {
      return JSON.parse(first);
    } catch (_err) {
      return null;
    }
  }
  return first;
}

export async function handleDyingOnUpdate(actorDocument, change) {
  if (!game.settings.get(MODULE.ID, "enableDyingAutomation")) return;
  if (!hasPrimaryHealthUpdate(actorDocument, change)) return;
  if (isWoundsVigorActive(actorDocument)) return;

  const newHp = getPrimaryHealthValue(actorDocument, change);
  const isDead = actorDocument.statuses?.has?.("dead");

  if (isDead || newHp >= 0) {
    if (actorDocument.statuses?.has?.("dying")) {
      await actorDocument.setCondition("dying", false);
    }
    if (newHp >= 0 && actorDocument.statuses?.has?.("stable")) {
      await actorDocument.setCondition("stable", false);
    }
    return;
  }

  if (actorDocument.statuses?.has?.("stable")) {
    if (actorDocument.statuses?.has?.("dying")) {
      await actorDocument.setCondition("dying", false);
    }
    return;
  }

  if (!actorDocument.statuses?.has?.("dying")) {
    await actorDocument.setCondition("dying", true);
  }

  if (!actorDocument.statuses?.has?.("unconscious")) {
    await actorDocument.setCondition("unconscious", true);
  }
}

export function handleDyingPreActionUse(action) {
  if (!game.settings.get(MODULE.ID, "enableDyingAutomation")) return;

  const handling = game.settings.get(MODULE.ID, "dyingHandling");
  if (!handling || handling === "disabled") return;

  const token = action?.token;
  const actor = token?.actor;
  if (!actor) return;
  if (!isDyingStateRelevant(actor)) return;
  if (actor.statuses?.has?.("stable")) return;
  if (!actor.statuses?.has?.("dying")) return;

  if (handling === "strict") {
    action.shared.reject = true;
    ui.notifications.info(game.i18n.format("NAS.conditions.main.DyingStrict", { name: token.name }));
  } else if (handling === "lenient") {
    ui.notifications.info(game.i18n.format("NAS.conditions.main.DyingLenient", { name: token.name }));
  }
}

export async function handleDyingCombatUpdate(combat, update) {
  if (!game.settings.get(MODULE.ID, "enableDyingAutomation")) return;
  if (!combat) return;
  if (update?.turn === undefined && update?.round === undefined) return;

  const combatant = combat.combatant;
  if (!combatant) return;
  const token = combatant.token?.object ?? canvas.tokens.get(combatant.tokenId);
  const actor = token?.actor;
  if (!actor) return;
  if (!isDyingStateRelevant(actor)) return;
  if (actor.statuses?.has?.("stable")) return;

  if (!actor.statuses?.has?.("dying")) {
    await actor.setCondition("dying", true);
  }
  if (!actor.statuses?.has?.("unconscious")) {
    await actor.setCondition("unconscious", true);
  }

  const hp = Number(actor.system?.attributes?.hp?.value ?? 0) || 0;
  const conScore = Number(actor.system?.abilities?.con?.total ?? actor.system?.abilities?.con?.value ?? 0) || 0;
  if (conScore > 0 && hp <= -conScore) {
    return;
  }

  const result = await actor.rollAbilityTest("con", {
    skipDialog: true,
    token: token?.document ?? actor.token,
    dc: 10,
    bonus: `${hp}[${game.i18n.localize("NAS.conditions.main.DyingHpPenalty")}]` 
  });

  const rollData = getRollDataFromResult(result);
  const total = Number(rollData?.total ?? 0);
  const isNat20 = Boolean(rollData?.isNat20);
  const success = isNat20 || total >= 10;

  if (success) {
    if (!actor.statuses?.has?.("stable")) {
      await actor.setCondition("stable", true);
    }
    if (actor.statuses?.has?.("dying")) {
      await actor.setCondition("dying", false);
    }
    return;
  }

  await actor.update({ "system.attributes.hp.value": hp - 1 });
}
