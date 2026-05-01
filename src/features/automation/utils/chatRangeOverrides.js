import { MODULE } from "../../../common/module.js";
import { elementFromHtmlLike } from "../../../common/foundryCompat.js";

let hoverOverrideAttached = false;
const HOVER_ACTION_OVERRIDES = new WeakMap();

function getAttackReachApi() {
  const reach = pf1?.canvas?.attackReach;
  if (reach?.render && (reach.clear || reach.clearHighlight)) {
    return {
      render: reach.render.bind(reach),
      clear: (reach.clear ?? reach.clearHighlight).bind(reach),
      supportsOptions: true,
    };
  }
  if (reach?.showAttackReach) {
    return {
      render: reach.showAttackReach.bind(reach),
      clear: (reach.clearHighlight ?? reach.clear)?.bind(reach),
      supportsOptions: false,
    };
  }
  const canvasApi = pf1?.canvas;
  if (canvasApi?.showAttackReach && canvasApi?.clearHighlight) {
    return {
      render: canvasApi.showAttackReach.bind(canvasApi),
      clear: canvasApi.clearHighlight.bind(canvasApi),
      supportsOptions: false,
    };
  }
  return null;
}

function normalizeRangeUnits(units, supportsOptions) {
  if (supportsOptions) return units;
  if (["close", "medium", "long"].includes(units)) return "ft";
  return units;
}

function applyActionGetRangeOverride(action, overrides) {
  if (!action || !overrides) return;
  if (HOVER_ACTION_OVERRIDES.has(action)) return;

  const base = Number(overrides.base ?? 0);
  const min = Number(overrides.min ?? 0);
  const increments = Number(overrides.increments ?? 1);
  const units = overrides.units ?? "";

  HOVER_ACTION_OVERRIDES.set(action, {
    rangeUnits: action.range?.units,
    rangeMinUnits: action.range?.minUnits,
    rangeMaxIncrements: action.range?.maxIncrements,
    rangeValue: action.range?.value,
    rangeMinValue: action.range?.minValue,
    touch: action.touch,
    sourceUnits: action._source?.range?.units,
    sourceMinUnits: action._source?.range?.minUnits,
    sourceMaxIncrements: action._source?.range?.maxIncrements,
    sourceValue: action._source?.range?.value,
    sourceMinValue: action._source?.range?.minValue,
    sourceTouch: action._source?.touch,
  });

  const update = {};
  if (units) {
    update["range.units"] = units;
    update["range.minUnits"] = units;
  }
  if (Number.isFinite(increments)) {
    update["range.maxIncrements"] = increments;
  }
  if (Number.isFinite(base)) {
    update["range.value"] = String(base);
  }
  if (Number.isFinite(min)) {
    update["range.minValue"] = String(min);
  }
  if (units && units !== "touch") {
    update.touch = false;
  }

  if (typeof action.updateSource === "function") {
    action.updateSource(update);
  } else if (action._source) {
    action._source.range ??= {};
    if ("range.units" in update) action._source.range.units = update["range.units"];
    if ("range.minUnits" in update) action._source.range.minUnits = update["range.minUnits"];
    if ("range.maxIncrements" in update) action._source.range.maxIncrements = update["range.maxIncrements"];
    if ("range.value" in update) action._source.range.value = update["range.value"];
    if ("range.minValue" in update) action._source.range.minValue = update["range.minValue"];
    if ("touch" in update) action._source.touch = update.touch;
  }

  if (action.range) {
    if ("range.units" in update) action.range.units = update["range.units"];
    if ("range.minUnits" in update) action.range.minUnits = update["range.minUnits"];
    if ("range.maxIncrements" in update) action.range.maxIncrements = update["range.maxIncrements"];
    if ("range.value" in update) action.range.value = update["range.value"];
    if ("range.minValue" in update) action.range.minValue = update["range.minValue"];
  }
  if (update.touch === false) {
    action.touch = false;
  }
}

function restoreActionGetRangeOverride(action) {
  if (!action || !HOVER_ACTION_OVERRIDES.has(action)) return;
  const original = HOVER_ACTION_OVERRIDES.get(action);
  HOVER_ACTION_OVERRIDES.delete(action);
  if (!original) return;

  if (action.range) {
    action.range.units = original.rangeUnits;
    action.range.minUnits = original.rangeMinUnits;
    action.range.maxIncrements = original.rangeMaxIncrements;
    action.range.value = original.rangeValue;
    action.range.minValue = original.rangeMinValue;
  }
  action.touch = original.touch;
  if (typeof action.updateSource === "function") {
    action.updateSource({
      "range.units": original.sourceUnits,
      "range.minUnits": original.sourceMinUnits,
      "range.maxIncrements": original.sourceMaxIncrements,
      "range.value": original.sourceValue,
      "range.minValue": original.sourceMinValue,
      touch: original.sourceTouch,
    });
  } else if (action._source) {
    action._source.range ??= {};
    action._source.range.units = original.sourceUnits;
    action._source.range.minUnits = original.sourceMinUnits;
    action._source.range.maxIncrements = original.sourceMaxIncrements;
    action._source.range.value = original.sourceValue;
    action._source.range.minValue = original.sourceMinValue;
    action._source.touch = original.sourceTouch;
  }

}

export function applyChatRangeOverrides(message, html) {
  const overrides = message?.flags?.[MODULE.ID]?.actionOverrides?.range;
  if (!overrides) return;

  const root = elementFromHtmlLike(html);
  if (!root) return;

  const rangeElement = root.querySelector(".card-range");
  if (!rangeElement) return;

  if (overrides.min != null) rangeElement.dataset.rangeMin = String(overrides.min);
  if (overrides.max != null) rangeElement.dataset.rangeMax = String(overrides.max);
  if (overrides.base != null) rangeElement.dataset.rangeBase = String(overrides.base);
  if (overrides.units) rangeElement.dataset.units = String(overrides.units);
  if (overrides.increments != null) rangeElement.dataset.increments = String(overrides.increments);
}

export function registerChatRangeHoverOverrides(html) {
  if (hoverOverrideAttached) return;
  const root = elementFromHtmlLike(html);
  if (!root) return;
  hoverOverrideAttached = true;

  root.addEventListener(
    "pointerenter",
    (event) => {
      const elem = event.target?.closest?.(".card-range");
      if (!elem) return;
      const messageId = elem.closest(".chat-message[data-message-id]")?.dataset?.messageId;
      const msg = messageId ? game.messages.get(messageId) : null;
      const action = msg?.actionSource;
      const overrides = msg?.flags?.[MODULE.ID]?.actionOverrides?.range;
      if (!overrides || !action) return;

      const { scene: sceneId, token: tokenId } = msg?.speaker ?? {};
      const actor = ChatMessage.implementation.getSpeakerActor(msg?.speaker);
      const tokenDoc = actor?.token ?? game.scenes.get(sceneId)?.tokens.get(tokenId);
      const token = tokenDoc?.object ?? tokenDoc;
      if (!token) return;

      const api = getAttackReachApi();
      if (!api?.render) return;

    const normalizedUnits = normalizeRangeUnits(overrides.units, api.supportsOptions);
    const normalizedOverrides = normalizedUnits ? { ...overrides, units: normalizedUnits } : overrides;
      applyActionGetRangeOverride(action, normalizedOverrides);
      api.render(token, action);
      event.stopImmediatePropagation();
    },
    { capture: true, passive: true }
  );

  root.addEventListener(
    "pointerleave",
    (event) => {
      const elem = event.target?.closest?.(".card-range");
      if (!elem) return;
      const api = getAttackReachApi();
      if (!api?.clear) return;
      const messageId = elem.closest(".chat-message[data-message-id]")?.dataset?.messageId;
      const msg = messageId ? game.messages.get(messageId) : null;
      const action = msg?.actionSource;
      restoreActionGetRangeOverride(action);
      api.clear();
    },
    { capture: true, passive: true }
  );
}
