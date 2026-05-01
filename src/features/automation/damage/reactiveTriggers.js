import { MODULE } from "../../../common/module.js";
import { chatMessageStyle } from "../../../common/foundryCompat.js";
import { toDamagePartModel } from "./instances.js";
import { applyReactiveDamageToActor, toggleReactiveBuff, toggleReactiveCondition } from "../../../integration/moduleSockets.js";
import { actorUsesWoundsVigor, isWvNoWoundsActor } from "../utils/woundsVigor.js";

const REACTIVE_FLAG_KEY = "itemReactiveEffects";

function localize(path) {
  return game.i18n.localize(`NAS.reactive.${path}`);
}

function escHtml(value) {
  return foundry.utils.escapeHTML(String(value ?? ""));
}

function displayNameForOwnerItem(ownerItem, kind) {
  const raw = String(ownerItem?.name ?? "").trim();
  if (raw) return raw;
  return kind === "onHit"
    ? game.i18n.localize("NAS.reactive.chatSummary.onHitFallback")
    : game.i18n.localize("NAS.reactive.chatSummary.onStruckFallback");
}

function resolveConditionDisplayName(conditionId) {
  const id = String(conditionId ?? "").trim();
  if (!id) return game.i18n.localize("NAS.reactive.chatSummary.unknownCondition");
  for (const c of pf1?.registry?.conditions ?? []) {
    if (String(c?._id ?? "") === id) return String(c?.name ?? id);
  }
  return id;
}

function resolveBuffDisplayName(buffUuid) {
  const uuid = String(buffUuid ?? "").trim();
  if (!uuid) return game.i18n.localize("NAS.reactive.chatSummary.unknownBuff");
  try {
    if (typeof foundry.utils?.fromUuidSync === "function") {
      const doc = foundry.utils.fromUuidSync(uuid);
      if (doc?.name) return String(doc.name);
    }
  } catch (_err) {
    return game.i18n.localize("NAS.reactive.chatSummary.unknownBuff");
  }
  return game.i18n.localize("NAS.reactive.chatSummary.unknownBuff");
}

function actorChatName(actor) {
  return actor?.name ?? game.i18n.localize("NAS.common.labels.target");
}

function formatDamageTypeNames(typeIds) {
  const ids = Array.isArray(typeIds) ? typeIds.map((id) => String(id ?? "").trim()).filter(Boolean) : [];
  if (!ids.length) return localize("damageTypeUntyped");
  const labels = ids.map((raw) => {
    const id = raw || "untyped";
    if (id === "untyped") return localize("damageTypeUntyped");
    for (const [, value] of pf1?.registry?.damageTypes?.entries?.() ?? []) {
      if (String(value?.id ?? "") === id) return String(value?.name ?? id);
    }
    return id;
  });
  return [...new Set(labels)].join(", ");
}

function getReactiveFlags(item) {
  return foundry.utils.deepClone(item?.flags?.[MODULE.ID]?.[REACTIVE_FLAG_KEY] ?? {});
}

function getActionIdFromOptions(options = {}) {
  return String(
    options?.action?.id
    ?? options?.actionId
    ?? options?.message?.system?.action?.id
    ?? ""
  ).trim();
}

function resolveActorFromSpeaker(speaker) {
  const actorId = speaker?.actor;
  if (actorId) {
    const actor = game.actors?.get?.(actorId);
    if (actor) return actor;
  }
  const tokenId = speaker?.token;
  if (tokenId) {
    const token = canvas?.tokens?.get?.(tokenId);
    const actor = token?.actor ?? null;
    if (actor) return actor;
  }
  return null;
}

function resolveActorFromNasReactiveAttackerUuid(options) {
  const uuid = options?._nasReactiveAttackerUuid;
  if (!uuid || typeof uuid !== "string") return null;
  try {
    if (typeof foundry.utils?.fromUuidSync === "function") {
      const doc = foundry.utils.fromUuidSync(uuid);
      if (doc && doc.documentName === "Actor") return doc;
    }
  } catch (_err) {
    return null;
  }
  return null;
}

export function resolveSourceActorFromOptions(options = {}) {
  const fromNasUuid = resolveActorFromNasReactiveAttackerUuid(options);
  if (fromNasUuid) return fromNasUuid;
  const direct =
    options?.action?.actor
    ?? options?.item?.actor
    ?? options?.message?.actor
    ?? null;
  if (direct) return direct;
  return resolveActorFromSpeaker(options?.message?.speaker);
}

function normalizeDamageTypes(effect = {}) {
  const fromArray = Array.isArray(effect?.damageTypes)
    ? effect.damageTypes.map((d) => String(d ?? "").trim()).filter(Boolean)
    : [];
  if (fromArray.length) return [...new Set(fromArray)];
  const one = String(effect?.damageType ?? "").trim();
  return one ? [one] : ["untyped"];
}

function normalizeEffect(effect = {}) {
  const damageTypes = normalizeDamageTypes(effect);
  return {
    type: String(effect?.type ?? ""),
    mode: String(effect?.mode ?? "formula"),
    value: Number(effect?.value) || 0,
    formula: String(effect?.formula ?? ""),
    damageType: String(effect?.damageType ?? damageTypes[0] ?? "untyped"),
    damageTypes,
    buffUuid: String(effect?.buffUuid ?? ""),
    conditionId: String(effect?.conditionId ?? ""),
    message: effect?.message !== false
  };
}

const TOGGLE_EFFECT_TYPES = new Set([
  "applyConditionAttacker",
  "removeConditionAttacker",
  "applyBuffAttacker",
  "removeBuffAttacker",
  "applyConditionTarget",
  "removeConditionTarget",
  "applyBuffTarget",
  "removeBuffTarget"
]);

function effectActsOnTarget(effect) {
  return String(effect?.type ?? "").endsWith("Target");
}

function buildReactiveLineHtml({ effect, magnitude, affectedActor, triggerKind: _triggerKind }) {
  void _triggerKind;
  const actor = escHtml(actorChatName(affectedActor));
  const effectType = String(effect?.type ?? "");
  const isToggleEffect = TOGGLE_EFFECT_TYPES.has(effectType);

  if (isToggleEffect) {
    if (effectType.includes("Condition")) {
      const name = escHtml(resolveConditionDisplayName(effect.conditionId));
      if (effectType.startsWith("apply")) {
        return game.i18n.format("NAS.reactive.chatSummary.lineApplyCondition", { name, actor });
      }
      return game.i18n.format("NAS.reactive.chatSummary.lineRemoveCondition", { name, actor });
    }
    const name = escHtml(resolveBuffDisplayName(effect.buffUuid));
    if (effectType.startsWith("apply")) {
      return game.i18n.format("NAS.reactive.chatSummary.lineApplyBuff", { name, actor });
    }
    return game.i18n.format("NAS.reactive.chatSummary.lineRemoveBuff", { name, actor });
  }

  if (effectType === "healAttacker") {
    return game.i18n.format("NAS.reactive.chatSummary.lineHeal", { actor, amount: magnitude });
  }
  if (effectType === "damageAttacker") {
    const types = escHtml(formatDamageTypeNames(effect?.damageTypes));
    return game.i18n.format("NAS.reactive.chatSummary.lineDamage", { actor, amount: magnitude, types });
  }

  return null;
}

function postReactiveChatSummary({ ownerActor, otherActor, ownerItemPlain, triggerKind, finalDamage, lineHtmls }) {
  if (!lineHtmls.length) return;

  const titleKey =
    triggerKind === "onHit" ? "NAS.reactive.chatSummary.onHitTitle" : "NAS.reactive.chatSummary.onStruckTitle";
  const title = game.i18n.format(titleKey, { item: escHtml(ownerItemPlain) });
  const subtitleKey =
    triggerKind === "onHit"
      ? "NAS.reactive.chatSummary.subtitleOnHit"
      : "NAS.reactive.chatSummary.subtitleOnStruck";
  const subtitle = game.i18n.format(subtitleKey, {
    other: escHtml(actorChatName(otherActor)),
    finalDamage: String(Math.max(0, Number(finalDamage) || 0))
  });

  const innerLines = lineHtmls.map((html) => `<li>${html}</li>`).join("");
  const content = [
    `<div class="nas-reactive-chat-summary" data-nas-reactive-summary>`,
    `<div class="nas-reactive-chat-header"><strong>${title}</strong></div>`,
    `<div class="nas-reactive-chat-subtitle">${subtitle}</div>`,
    `<ul class="nas-reactive-chat-lines">`,
    innerLines,
    `</ul></div>`
  ].join("");

  const speakerData = ChatMessage.getSpeaker({ actor: ownerActor ?? null });
  ChatMessage.create({
    ...chatMessageStyle("OTHER"),
    user: game.user?.id,
    speaker: speakerData,
    content
  });
}

function buildOnHitFromRaw(raw) {
  if (!raw || raw.enabled !== true || !Array.isArray(raw.effects) || raw.effects.length === 0) return null;
  return {
    effects: raw.effects.map(normalizeEffect)
  };
}

function getOnHitConfig(sourceItem, actionId) {
  const flags = getReactiveFlags(sourceItem);
  const overrideRaw = flags?.onHitByActionOverride?.[actionId];
  const fromOverride = buildOnHitFromRaw(overrideRaw);
  if (fromOverride) return fromOverride;
  const baseRaw = flags?.onHitByAction?.[actionId] ?? null;
  return buildOnHitFromRaw(baseRaw);
}

function isItemOnStruckActive(item) {
  if (!item) return false;
  if (item.type === "buff") return item.system?.active === true;
  if (item.type === "equipment") return item.system?.equipped === true;
  return false;
}

function getOnStruckConfigs(targetActor) {
  if (!targetActor?.items) return [];
  const out = [];
  for (const item of targetActor.items) {
    if (!["buff", "equipment"].includes(item?.type)) continue;
    if (!isItemOnStruckActive(item)) continue;
    const flags = getReactiveFlags(item);
    const raw = flags?.onStruck ?? null;
    if (!raw || raw.enabled !== true || !Array.isArray(raw.effects) || raw.effects.length === 0) continue;
    out.push({
      sourceItem: item,
      filters: {
        meleeOnly: raw?.filters?.meleeOnly !== false,
        excludeReach: raw?.filters?.excludeReach !== false
      },
      effects: raw.effects.map(normalizeEffect)
    });
  }
  return out;
}

function resolveActionType(options = {}) {
  return String(
    options?.action?.actionType
    ?? options?.message?.system?.action?.actionType
    ?? ""
  ).toLowerCase();
}

const _PF1_MELEE_LIKE = new Set(["mwak", "msak", "mcman"]);

function resolveRangeUnits(options = {}) {
  return String(
    options?.action?.range?.units ?? options?.message?.system?.action?.range?.units ?? ""
  ).toLowerCase();
}

function isMeleeContext(options = {}) {
  const type = resolveActionType(options);
  return _PF1_MELEE_LIKE.has(type);
}

function isReachContext(options = {}) {
  return resolveRangeUnits(options) === "reach";
}

function buildRollData({ sourceActor, targetActor, finalDamage = 0 }) {
  const sourceRollData = sourceActor?.getRollData?.() ?? {};
  const targetRollData = targetActor?.getRollData?.() ?? {};
  return {
    ...targetRollData,
    ...sourceRollData,
    nas: {
      finalDamage: Number(finalDamage) || 0
    },
    attacker: sourceRollData,
    target: targetRollData
  };
}

async function evaluateMagnitude(effect, context) {
  const mode = String(effect?.mode ?? "formula");
  if (mode === "percentOfFinalDamage") {
    const pct = Number(effect?.value) || 0;
    const base = Math.max(0, Number(context?.finalDamage) || 0);
    const out = Math.floor((base * pct) / 100);
    return out;
  }
  if (mode === "formula") {
    const formula = String(effect?.formula ?? "").trim();
    if (!formula) return 0;
    try {
      const roll = await new Roll(formula, context.rollData).evaluate();
      return Math.max(0, Math.floor(Number(roll?.total) || 0));
    } catch (_err) {
      return 0;
    }
  }
  return 0;
}

function buildDamageInstances(amount, effect) {
  const typesRaw = normalizeDamageTypes(effect);
  const types = typesRaw.length ? typesRaw : ["untyped"];
  return [toDamagePartModel({ types, value: amount, formula: String(amount) })];
}

function getHealthSnapshotForDelta(actor) {
  const a = actor?.system?.attributes;
  if (!a) return null;
  if (actorUsesWoundsVigor(actor)) {
    const noW = isWvNoWoundsActor(actor);
    const vigor = (Number(a.vigor?.value) || 0) + (Number(a.vigor?.temp) || 0);
    const wounds = noW || a.wounds?.value === undefined ? 0 : Number(a.wounds?.value) || 0;
    return { kind: "wv", vigor, wounds, noWounds: noW };
  }
  if (a.hp && a.hp.value !== undefined) {
    return { kind: "hp", hp: Number(a.hp.value) || 0 };
  }
  return null;
}

function hurtDeltaFromSnapshots(pre, post, isHeal) {
  if (!pre || !post) return null;
  if (pre.kind !== post.kind) return null;
  if (pre.kind === "hp" && post.kind === "hp") {
    if (isHeal) return Math.max(0, post.hp - pre.hp);
    return Math.max(0, pre.hp - post.hp);
  }
  if (pre.kind === "wv" && post.kind === "wv") {
    if (isHeal) {
      const hv = Math.max(0, post.vigor - pre.vigor);
      const hW = pre.noWounds ? 0 : Math.max(0, post.wounds - pre.wounds);
      return hv + hW;
    }
    const dV = Math.max(0, pre.vigor - post.vigor);
    const dW = pre.noWounds ? 0 : Math.max(0, pre.wounds - post.wounds);
    return dV + dW;
  }
  return null;
}

async function applyToActor(actor, effect, magnitude, metadata = {}) {
  if (!actor) return undefined;
  if (effect?.type === "applyConditionAttacker" || effect?.type === "applyConditionTarget") {
    if (!effect.conditionId) return;
    await toggleReactiveCondition(actor, effect.conditionId, true);
    return;
  }
  if (effect?.type === "removeConditionAttacker" || effect?.type === "removeConditionTarget") {
    if (!effect.conditionId) return;
    await toggleReactiveCondition(actor, effect.conditionId, false);
    return;
  }
  if (effect?.type === "applyBuffAttacker" || effect?.type === "applyBuffTarget") {
    if (!effect.buffUuid) return;
    await toggleReactiveBuff(actor, effect.buffUuid, true);
    return;
  }
  if (effect?.type === "removeBuffAttacker" || effect?.type === "removeBuffTarget") {
    if (!effect.buffUuid) return;
    await toggleReactiveBuff(actor, effect.buffUuid, false);
    return;
  }
  if (!Number.isFinite(magnitude) || magnitude <= 0) return undefined;
  const isHeal = effect.type === "healAttacker";
  const value = isHeal ? -Math.abs(magnitude) : Math.abs(magnitude);
  const options = {
    dialog: false,
    _nasReactiveEffect: true,
    _nasReactiveSource: metadata.source ?? "unknown"
  };
  if (metadata.attackerUuid) {
    options._nasReactiveAttackerUuid = metadata.attackerUuid;
  }
  if (!isHeal) {
    options.instances = buildDamageInstances(Math.abs(magnitude), effect);
  }
  const pre = getHealthSnapshotForDelta(actor);
  await applyReactiveDamageToActor(actor, value, options);
  const post = getHealthSnapshotForDelta(actor);
  const hurtDelta = hurtDeltaFromSnapshots(pre, post, isHeal);
  if (String(effect.type) === "healAttacker" || String(effect.type) === "damageAttacker") {
    if (hurtDelta != null) return hurtDelta;
  }
  return undefined;
}

async function applyEffects({ sourceActor, targetActor, effects, finalDamage, options, triggerMeta }) {
  void options;
  const { kind, ownerActor, otherActor, ownerItem, sourceTag } = triggerMeta;
  const ownerItemPlain = displayNameForOwnerItem(ownerItem, kind);
  const lineHtmls = [];
  const rollData = buildRollData({ sourceActor, targetActor, finalDamage });

  for (const effect of effects) {
    const isToggleEffect = TOGGLE_EFFECT_TYPES.has(effect?.type);
    const magnitude = isToggleEffect ? 1 : await evaluateMagnitude(effect, { finalDamage, rollData });
    if (!isToggleEffect && magnitude <= 0) continue;
    const affectedActor = effectActsOnTarget(effect) ? targetActor : sourceActor;
    const effType = String(effect.type ?? "");
    const actualApplied = await applyToActor(affectedActor, effect, magnitude, { source: sourceTag, attackerUuid: sourceActor?.uuid ?? null });
    const lineMagnitude = (effType === "healAttacker" || effType === "damageAttacker") && Number.isFinite(actualApplied) ? actualApplied : magnitude;

    if (effect.message !== true) continue;
    if (!isToggleEffect && magnitude <= 0) continue;
    const line = buildReactiveLineHtml({
      effect,
      magnitude: lineMagnitude,
      affectedActor,
      triggerKind: kind
    });
    if (line) lineHtmls.push(line);
  }

  if (lineHtmls.length) {
    postReactiveChatSummary({
      ownerActor,
      otherActor,
      ownerItemPlain,
      triggerKind: kind,
      finalDamage,
      lineHtmls
    });
  }
}

export async function applyReactiveEffectsForHit({
  sourceActor,
  sourceItem,
  targetActor,
  options = {},
  finalDamage = 0
} = {}) {
  if (!sourceActor || !targetActor) return;
  if (options?._nasReactiveEffect) return;

  const actionId = getActionIdFromOptions(options);
  const onHit = getOnHitConfig(sourceItem, actionId);
  if (onHit) {
    await applyEffects({
      sourceActor,
      targetActor,
      effects: onHit.effects,
      finalDamage,
      options,
      triggerMeta: {
        kind: "onHit",
        ownerActor: sourceActor,
        otherActor: targetActor,
        ownerItem: sourceItem,
        sourceTag: "onHit"
      }
    });
  }

  const meleeContext = isMeleeContext(options);
  const reachContext = isReachContext(options);
  const onStruckConfigs = getOnStruckConfigs(targetActor);
  for (const entry of onStruckConfigs) {
    const filters = entry.filters ?? {};
    if (filters.meleeOnly && !meleeContext) continue;
    if (filters.excludeReach && reachContext) continue;
    await applyEffects({
      sourceActor,
      targetActor,
      effects: entry.effects,
      finalDamage,
      options,
      triggerMeta: {
        kind: "onStruck",
        ownerActor: targetActor,
        otherActor: sourceActor,
        ownerItem: entry.sourceItem,
        sourceTag: "onStruck"
      }
    });
  }
}
