import { MODULE } from "../../../common/module.js";
import { consumeCombatTextContext } from "./combatTextContext.js";

let hookRegistered = false;

const COMBAT_TEXT_MODES = new Set(["off", "enhanced", "cinematic"]);
const CINEMATIC_PRESETS = new Set(["arcLanes", "tokenSides"]);

const CATEGORY_STYLES = {
  hp: {
    positive: "#35ff7a",
    negative: "#ff3b30",
    direction: { positive: "TOP", negative: "BOTTOM" },
    jitter: 0.65,
    distance: 0.75,
    laneOffset: { x: 0.22, y: 0 },
    tokenSide: "right",
    arc: 0.2,
    strokeThickness: 3
  },
  vigor: {
    positive: "#35ff7a",
    negative: "#ff3b30",
    direction: { positive: "TOP", negative: "BOTTOM" },
    jitter: 0.65,
    distance: 0.75,
    laneOffset: { x: 0.12, y: -0.12 },
    tokenSide: "right",
    arc: -0.2,
    strokeThickness: 3
  },
  wounds: {
    positive: "#46ff8a",
    negative: "#ff1f4f",
    direction: { positive: "TOP", negative: "BOTTOM" },
    jitter: 0.75,
    distance: 0.8,
    laneOffset: { x: 0.34, y: 0.08 },
    tokenSide: "right",
    arc: 0.25,
    strokeThickness: 3.25
  },
  temp: {
    positive: "#8cff2f",
    negative: "#ff8a2a",
    direction: { positive: "TOP", negative: "RIGHT" },
    jitter: 0.85,
    distance: 0.55,
    laneOffset: { x: 0.28, y: -0.24 },
    tokenSide: "right",
    arc: -0.35,
    strokeThickness: 2.5
  },
  nonlethal: {
    positive: "#b44cff",
    negative: "#21d36b",
    direction: { positive: "LEFT", negative: "RIGHT" },
    jitter: 0.9,
    distance: 0.55,
    laneOffset: { x: 0, y: 0.24 },
    tokenSide: "left",
    arc: 0.3,
    strokeThickness: 2.75
  },
  ability: {
    positive: "#5cffc8",
    negative: "#ff9b2f",
    direction: { positive: "TOP", negative: "BOTTOM" },
    jitter: 0.65,
    distance: 0.45,
    laneOffset: { x: -0.34, y: 0 },
    tokenSide: "left",
    arc: -0.45,
    strokeThickness: 3
  },
  energyDrain: {
    positive: "#4dff9b",
    negative: "#b16dff",
    direction: { positive: "TOP", negative: "BOTTOM" },
    jitter: 0.95,
    distance: 0.65,
    laneOffset: { x: -0.22, y: -0.12 },
    tokenSide: "left",
    arc: -0.3,
    strokeThickness: 3
  }
};

function settingRegistered(settingName) {
  return globalThis.game?.settings?.settings?.has?.(`${MODULE.ID}.${settingName}`) === true;
}

function combatTextMode() {
  if (settingRegistered("combatTextMode")) {
    const mode = game.settings.get(MODULE.ID, "combatTextMode");
    if (COMBAT_TEXT_MODES.has(mode)) return mode;
  }

  // Compatibility for worlds that enabled the first boolean version before modes existed.
  if (settingRegistered("enhancedCombatText") && game.settings.get(MODULE.ID, "enhancedCombatText") === true) {
    return "enhanced";
  }

  return "off";
}

function cinematicCombatTextPreset() {
  if (!settingRegistered("cinematicCombatTextPreset")) return "arcLanes";
  const preset = game.settings.get(MODULE.ID, "cinematicCombatTextPreset");
  return CINEMATIC_PRESETS.has(preset) ? preset : "arcLanes";
}

function categoryForDeltaKey(key = "") {
  const parts = String(key).split(".");
  const [root, branch] = parts;
  if (root === "hp" && branch === "temp") return "temp";
  if (root === "vigor" && branch === "temp") return "temp";
  if (root === "hp" && branch === "nonlethal") return "nonlethal";
  if (root === "energyDrain") return "energyDrain";
  if (["hp", "vigor", "wounds"].includes(root)) return root;
  if (parts.some((part) => ["damage", "drain", "userPenalty", "penalty"].includes(part))) return "ability";
  return null;
}

function directionPoint(name, fallback, current) {
  return globalThis.CONST?.TEXT_ANCHOR_POINTS?.[name] ?? fallback ?? current;
}

function magnitudeFontBonus(value) {
  const amount = Math.abs(Number(value) || 0);
  if (amount >= 50) return 10;
  if (amount >= 25) return 7;
  if (amount >= 10) return 4;
  if (amount >= 5) return 2;
  return 0;
}

function scrollDistance(style, current) {
  const gridSize = globalThis.canvas?.grid?.size ?? 72;
  const preferred = Math.max(gridSize * style.distance, 24);
  const existing = Number(current);
  return Number.isFinite(existing) && existing > 0 ? Math.min(existing, preferred) : preferred;
}

function laneOffset(style) {
  const gridSize = globalThis.canvas?.grid?.size ?? 72;
  return {
    x: gridSize * (Number(style?.laneOffset?.x) || 0),
    y: gridSize * (Number(style?.laneOffset?.y) || 0)
  };
}

function applyEnhancedTextData(options = {}, textData = {}, style) {
  const positive = options.positive === true;
  const directionName = positive ? style.direction.positive : style.direction.negative;
  const fallbackDirection = positive
    ? globalThis.CONST?.TEXT_ANCHOR_POINTS?.TOP
    : globalThis.CONST?.TEXT_ANCHOR_POINTS?.BOTTOM;
  const baseFontSize = Number(textData.fontSize) || Math.max((globalThis.canvas?.grid?.size ?? 72) / 3, 24);

  textData.fill = positive ? style.positive : style.negative;
  textData.direction = directionPoint(directionName, fallbackDirection, textData.direction);
  textData.stroke = 0x000000;
  textData.strokeThickness = style.strokeThickness;
  textData.jitter = Math.max(Number(textData.jitter) || 0, style.jitter);
  textData.distance = scrollDistance(style, textData.distance);
  textData.fontSize = baseFontSize + magnitudeFontBonus(options.value);
}

function canRenderCinematicText() {
  return Boolean(
    globalThis.canvas?.interface?.addChild
      && globalThis.PIXI?.Text
      && globalThis.requestAnimationFrame
  );
}

function resolveActorToken(actor) {
  return actor?.token?.object
    ?? actor?.token
    ?? actor?.getActiveTokens?.(true, true)?.[0]
    ?? actor?.getActiveTokens?.()?.[0]
    ?? null;
}

function textAnchorPoint(anchor) {
  return {
    [globalThis.CONST?.TEXT_ANCHOR_POINTS?.CENTER]: [0.5, 0.5],
    [globalThis.CONST?.TEXT_ANCHOR_POINTS?.BOTTOM]: [0.5, 0],
    [globalThis.CONST?.TEXT_ANCHOR_POINTS?.TOP]: [0.5, 1],
    [globalThis.CONST?.TEXT_ANCHOR_POINTS?.LEFT]: [1, 0.5],
    [globalThis.CONST?.TEXT_ANCHOR_POINTS?.RIGHT]: [0, 0.5]
  }[anchor ?? globalThis.CONST?.TEXT_ANCHOR_POINTS?.CENTER] ?? [0.5, 0.5];
}

function directionVector(direction) {
  switch (direction) {
    case globalThis.CONST?.TEXT_ANCHOR_POINTS?.BOTTOM:
      return { x: 0, y: 1 };
    case globalThis.CONST?.TEXT_ANCHOR_POINTS?.LEFT:
      return { x: -1, y: 0 };
    case globalThis.CONST?.TEXT_ANCHOR_POINTS?.RIGHT:
      return { x: 1, y: 0 };
    case globalThis.CONST?.TEXT_ANCHOR_POINTS?.TOP:
    default:
      return { x: 0, y: -1 };
  }
}

function easeOutCubic(t) {
  return 1 - Math.pow(1 - t, 3);
}

function tokenDimensions(token) {
  const gridSize = globalThis.canvas?.grid?.size ?? 72;
  return {
    width: Number(token?.w ?? token?.width ?? token?.object?.w) || gridSize,
    height: Number(token?.h ?? token?.height ?? token?.object?.h) || gridSize
  };
}

function tokenSidePosition({ origin, token, style, positive, travel, eased, t }) {
  const dimensions = tokenDimensions(token);
  const gridSize = globalThis.canvas?.grid?.size ?? 72;
  const sideSign = style.tokenSide === "left" ? -1 : 1;
  const gap = Math.max(gridSize * 0.1, 8);
  const curveRadiusX = Math.max(gridSize * 0.28, dimensions.width * 0.44, 18);
  const curveRadiusY = Math.max(travel * 0.95, dimensions.height * 0.88, gridSize * 0.55);
  const laneY = gridSize * (Number(style?.laneOffset?.y) || 0);
  const topAngle = -Math.PI / 2;
  const bottomAngle = Math.PI / 2;
  const angle = positive
    ? bottomAngle + ((topAngle - bottomAngle) * eased)
    : topAngle + ((bottomAngle - topAngle) * eased);
  const sideDistance = (dimensions.width / 2) + gap + (Math.cos(angle) * curveRadiusX);

  return {
    x: origin.x + (sideSign * sideDistance),
    y: origin.y + laneY + (Math.sin(angle) * curveRadiusY)
  };
}

function arcLanePosition({ start, vector, perpendicular, travel, arc, eased, t }) {
  const arcOffset = Math.sin(Math.PI * t) * arc;
  return {
    x: start.x + (vector.x * travel * eased) + (perpendicular.x * arcOffset),
    y: start.y + (vector.y * travel * eased) + (perpendicular.y * arcOffset)
  };
}

function renderCinematicTextInternal(actor, options = {}, textData = {}, style, context) {
  if (!canRenderCinematicText()) return false;

  try {
    if (game.settings.get("core", "scrollingStatusText") !== true) return true;
  } catch (_err) {
    // If the core setting is unavailable, let the renderer continue.
  }

  const token = resolveActorToken(actor);
  const origin = token?.center ?? token?.object?.center;
  if (!origin) return false;

  const TextClass = globalThis.foundry?.canvas?.containers?.PreciseText ?? globalThis.PIXI.Text;
  const styleInput = {
    ...textData,
    strokeThickness: context?.isCritical ? Math.max(textData.strokeThickness ?? 0, 5) : textData.strokeThickness,
    fontSize: context?.isCritical ? Math.round((Number(textData.fontSize) || 24) * 1.35) : textData.fontSize
  };
  const { direction, distance, jitter, duration, anchor, ...textStyle } = styleInput;
  const pixiStyle = typeof TextClass.getTextStyle === "function"
    ? TextClass.getTextStyle({ anchor, ...textStyle })
    : new globalThis.PIXI.TextStyle(textStyle);
  const text = new TextClass(options.label ?? "", pixiStyle);
  const layer = globalThis.canvas.interface;
  const scale = globalThis.canvas?.dimensions?.uiScale ?? 1;
  const travel = scrollDistance(style, distance);
  const vector = directionVector(direction);
  const perpendicular = { x: -vector.y, y: vector.x };
  const arc = travel * (style.arc ?? 0);
  const lane = laneOffset(style);
  const jitterX = ((Math.random() - 0.5) * (jitter ?? 0)) * text.width * scale;
  const jitterY = ((Math.random() - 0.5) * (jitter ?? 0)) * text.height * scale;
  const start = { x: origin.x + lane.x + jitterX, y: origin.y + lane.y + jitterY };
  const preset = cinematicCombatTextPreset();
  const baseDuration = Number(duration) || 1900;
  const critDuration = 2300;
  const presetDurationMultiplier = preset === "tokenSides" ? 1.75 : 1;
  const totalDuration = (context?.isCritical ? critDuration : baseDuration) * presetDurationMultiplier;
  const shake = context?.isCritical ? Math.max(4, (globalThis.canvas?.grid?.size ?? 72) * 0.06) : 0;
  const startedAt = performance.now();

  text.eventMode = "none";
  text.interactiveChildren = false;
  text.visible = true;
  text.alpha = 0;
  text.zIndex = globalThis.CONFIG?.Canvas?.groups?.interface?.zIndexScrollingText ?? 1100;
  text.anchor?.set?.(...textAnchorPoint(anchor));
  text.position.set(start.x, start.y);
  text.scale.set(scale * 0.65);
  layer.sortableChildren = true;
  layer.addChild(text);

  const animate = (time) => {
    const t = Math.min(1, (time - startedAt) / totalDuration);
    const eased = easeOutCubic(t);
    const fadeIn = Math.min(1, t / 0.18);
    const fadeOut = t > 0.62 ? Math.max(0, 1 - ((t - 0.62) / 0.38)) : 1;
    const pathPosition = preset === "tokenSides"
      ? tokenSidePosition({ origin, token, style, positive: options.positive === true, travel, eased, t })
      : arcLanePosition({ start, vector, perpendicular, travel, arc, eased, t });
    const shakeAmount = t < 0.35 ? shake * (1 - (t / 0.35)) : 0;
    const shakeX = shakeAmount ? (Math.random() - 0.5) * shakeAmount : 0;
    const shakeY = shakeAmount ? (Math.random() - 0.5) * shakeAmount : 0;
    const punch = context?.isCritical ? 1 + (Math.sin(Math.PI * Math.min(t / 0.3, 1)) * 0.22) : 1;

    text.alpha = fadeIn * fadeOut;
    text.position.set(
      pathPosition.x + shakeX,
      pathPosition.y + (preset === "tokenSides" ? jitterY : 0) + shakeY
    );
    text.scale.set(scale * punch);

    if (t < 1) {
      requestAnimationFrame(animate);
      return;
    }

    layer.removeChild(text);
    text.destroy();
  };

  requestAnimationFrame(animate);
  return true;
}

function renderCinematicText(actor, options = {}, textData = {}, style, context) {
  try {
    return renderCinematicTextInternal(actor, options, textData, style, context);
  } catch (err) {
    console.warn(`[${MODULE.ID}] Cinematic combat text failed; falling back to PF1 scrolling text.`, err);
    return false;
  }
}

function styleHealthDeltaText(actor, options = {}, textData = {}) {
  const mode = combatTextMode();
  if (mode === "off") return;

  const category = categoryForDeltaKey(options.key);
  const style = category ? CATEGORY_STYLES[category] : null;
  if (!style) return;

  applyEnhancedTextData(options, textData, style);

  if (mode !== "cinematic" || !canRenderCinematicText()) return;

  const context = consumeCombatTextContext(actor);
  if (renderCinematicText(actor, options, textData, style, context)) return false;
}

export function registerHealthDeltaTextEnhancer() {
  if (hookRegistered) return;
  hookRegistered = true;
  Hooks.on("pf1HealthDeltaRender", styleHealthDeltaText);
}
