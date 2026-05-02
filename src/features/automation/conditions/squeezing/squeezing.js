import { MODULE } from '../../../../common/module.js';
import { elementFromHtmlLike } from '../../../../common/foundryCompat.js';
import {
  HALF_SCALE,
  QUARTER_SCALE,
  computeNarrowMovementCost,
  computePassageWidthMetrics,
  computePathStateAnalysis,
  getTokenCenter,
  isPathReachable
} from '../../utils/squeezeRayMath.js';

const EXIT_HANDLING_IGNORE = "ignore";
const INTERNAL_MOVE_OPTION = "_nasSqueezeInternalMove";
const ESCAPE_ARTIST_SKILL = "esc";
const DEFAULT_ESCAPE_ARTIST_DC = 30;
const DEFAULT_MEDIUM_OVERALL_WIDTH = 1;
const DEFAULT_MEDIUM_HEAD_WIDTH = 0.25;
const ESCAPE_FAIL_STOP_BEFORE = "stopBeforeNarrow";
const ESCAPE_FAIL_ENTER_FIRST = "enterFirstNarrowSquare";
const TOKEN_FLAG_BODY_WIDTH = "squeezingBodyWidth";
const TOKEN_FLAG_HEAD_WIDTH = "squeezingHeadWidth";
let squeezingTokenConfigRegistered = false;

function hasMovementUpdate(updateData) {
  return Number.isFinite(updateData?.x) || Number.isFinite(updateData?.y);
}

function localizeOrFallback(key, fallback) {
  const localized = game.i18n?.localize?.(key);
  return localized && localized !== key ? localized : fallback;
}

export function registerSqueezingTokenConfigFields() {
  if (squeezingTokenConfigRegistered) return;
  squeezingTokenConfigRegistered = true;

  Hooks.on("renderTokenConfig", (app, html) => {
    const root = elementFromHtmlLike(html);
    const tokenDocument = app?.document ?? app?.object;
    if (!tokenDocument) return;

    const tabEl = root?.querySelector?.('.tab[data-tab="appearance"]');
    if (!tabEl) return;
    if (tabEl.querySelector?.('[data-nas-squeezing-token-fields="true"]')) return;

    const defaultBodyWidth = Number(game.settings.get(MODULE.ID, "squeezingMediumBodyWidth")) || DEFAULT_MEDIUM_OVERALL_WIDTH;
    const defaultHeadWidth = Number(game.settings.get(MODULE.ID, "squeezingMediumHeadWidth")) || DEFAULT_MEDIUM_HEAD_WIDTH;

    const bodyOverride = tokenDocument.getFlag(MODULE.ID, TOKEN_FLAG_BODY_WIDTH);
    const headOverride = tokenDocument.getFlag(MODULE.ID, TOKEN_FLAG_HEAD_WIDTH);
    const bodyOverrideValue = Number.isFinite(Number(bodyOverride)) ? Number(bodyOverride) : "";
    const headOverrideValue = Number.isFinite(Number(headOverride)) ? Number(headOverride) : "";

    const section = document.createElement("fieldset");
    section.dataset.nasSqueezingTokenFields = "true";
    section.classList.add("form-group", "stacked");
    section.innerHTML = `
      <legend>${localizeOrFallback("NAS.forms.squeezingTokenConfig.title", "Nevela Automation Suite - Squeezing")}</legend>
      <div class="form-group">
        <label>
          ${localizeOrFallback("NAS.forms.squeezingTokenConfig.labels.bodyWidth", "Body Width")}
          <a class="fa-solid fa-circle-info" data-tooltip="${localizeOrFallback("NAS.forms.squeezingTokenConfig.tooltips.bodyWidth", "Medium baseline body width. Leave empty to use module default.")}"></a>
        </label>
        <div class="form-fields">
          <input type="number" name="flags.${MODULE.ID}.${TOKEN_FLAG_BODY_WIDTH}" value="${bodyOverrideValue}" min="0" step="0.05" placeholder="${defaultBodyWidth}" />
        </div>
      </div>
      <div class="form-group">
        <label>
          ${localizeOrFallback("NAS.forms.squeezingTokenConfig.labels.headWidth", "Head Width")}
          <a class="fa-solid fa-circle-info" data-tooltip="${localizeOrFallback("NAS.forms.squeezingTokenConfig.tooltips.headWidth", "Medium baseline head width. Leave empty to use module default.")}"></a>
        </label>
        <div class="form-fields">
          <input type="number" name="flags.${MODULE.ID}.${TOKEN_FLAG_HEAD_WIDTH}" value="${headOverrideValue}" min="0" step="0.05" placeholder="${defaultHeadWidth}" />
        </div>
      </div>
    `;

    const shapeGroup = tabEl.querySelector?.('[name="shape"]')?.closest?.(".form-group");
    if (shapeGroup?.parentElement) {
      shapeGroup.parentElement.insertBefore(section, shapeGroup);
    } else {
      tabEl.appendChild(section);
    }
  });
}

function getNumericOrDefault(value, fallback, min = 0) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric) || numeric <= min) return fallback;
  return numeric;
}

function getPf1SubSmallSizeMultiplier(tokenDocument) {
  const sizeRaw = tokenDocument?.actor?.system?.traits?.size;
  const sizeKey = typeof sizeRaw === "string"
    ? sizeRaw
    : (
      sizeRaw?.base
      ?? sizeRaw?.size
      ?? sizeRaw?.key
      ?? sizeRaw?.value
      ?? sizeRaw?.label
      ?? tokenDocument?.actor?.system?.traits?.actualSize
    );
  const size = String(sizeKey ?? "").toLowerCase().trim();
  if (!size) return null;

  const pf1SpaceFactors = {
    f: 0.1,
    fine: 0.1,
    dim: 0.2,
    diminutive: 0.2,
    t: 0.5,
    tiny: 0.5,
    sm: 1,
    s: 1,
    small: 1,
    m: 1,
    med: 1,
    medium: 1,
    l: 2,
    lg: 2,
    large: 2,
    h: 3,
    huge: 3,
    g: 4,
    grg: 4,
    gargantuan: 4,
    c: 6,
    col: 6,
    colossal: 6
  };
  return pf1SpaceFactors[size] ?? null;
}

function getSqueezingGeometry(tokenDocument) {
  const gridSize = canvas?.grid?.size ?? 100;
  const pf1SizeMultiplier = getPf1SubSmallSizeMultiplier(tokenDocument);
  const tokenSizeMultiplier = getNumericOrDefault(
    pf1SizeMultiplier ?? tokenDocument.width,
    1,
    0
  );

  const moduleBodyWidth = getNumericOrDefault(
    game.settings.get(MODULE.ID, "squeezingMediumBodyWidth"),
    DEFAULT_MEDIUM_OVERALL_WIDTH,
    0
  );
  const moduleHeadWidth = getNumericOrDefault(
    game.settings.get(MODULE.ID, "squeezingMediumHeadWidth"),
    DEFAULT_MEDIUM_HEAD_WIDTH,
    0
  );
  const tokenBodyWidthOverride = getNumericOrDefault(
    tokenDocument.getFlag(MODULE.ID, TOKEN_FLAG_BODY_WIDTH),
    moduleBodyWidth,
    0
  );
  const tokenHeadWidthOverride = getNumericOrDefault(
    tokenDocument.getFlag(MODULE.ID, TOKEN_FLAG_HEAD_WIDTH),
    moduleHeadWidth,
    0
  );
  const escapeArtistDc = getNumericOrDefault(
    game.settings.get(MODULE.ID, "squeezingEscapeArtistDC"),
    DEFAULT_ESCAPE_ARTIST_DC,
    -1
  );

  // TODO: Support actor-level width overrides in addition to token overrides.
  const tokenOverallWidthPx = tokenBodyWidthOverride * tokenSizeMultiplier * gridSize;
  const tokenHeadWidthPx = tokenHeadWidthOverride * tokenSizeMultiplier * gridSize;

  return {
    tokenOverallWidthPx,
    tokenHeadWidthPx,
    halfOverallWidthPx: tokenOverallWidthPx * HALF_SCALE,
    escapeArtistDc
  };
}

function getPointAlongPath(originX, originY, destinationX, destinationY, t) {
  return {
    x: originX + ((destinationX - originX) * t),
    y: originY + ((destinationY - originY) * t)
  };
}

function getSnappedPoint(tokenDocument, x, y) {
  if (typeof tokenDocument.getSnappedPosition === "function") {
    const snapped = tokenDocument.getSnappedPosition({ x, y, elevation: tokenDocument.elevation ?? 0 });
    if (snapped && Number.isFinite(snapped.x) && Number.isFinite(snapped.y)) {
      return { x: snapped.x, y: snapped.y };
    }
  }
  return { x: Math.round(x), y: Math.round(y) };
}

function getEscapeFailurePoint(tokenDocument, originX, originY, destinationX, destinationY, firstSqueezedT, sampleSteps, failureHandling) {
  const stepT = sampleSteps > 0 ? 1 / sampleSteps : 0.01;
  let targetT = firstSqueezedT ?? 0;
  if (failureHandling === ESCAPE_FAIL_ENTER_FIRST) {
    targetT = Math.min(1, targetT + stepT);
  } else if (failureHandling === ESCAPE_FAIL_STOP_BEFORE) {
    targetT = Math.max(0, targetT - stepT);
  } else {
    targetT = Math.max(0, targetT - stepT);
  }
  let point = getPointAlongPath(originX, originY, destinationX, destinationY, targetT);
  let snappedPoint = getSnappedPoint(tokenDocument, point.x, point.y);
  let backstepIterations = 0;

  // For stop-before handling, continue stepping backward if snapping still lands on destination.
  if (failureHandling === ESCAPE_FAIL_STOP_BEFORE) {
    const maxBacksteps = Math.max(1, sampleSteps + 2);
    while (
      backstepIterations < maxBacksteps &&
      snappedPoint.x === destinationX &&
      snappedPoint.y === destinationY &&
      targetT > 0
    ) {
      targetT = Math.max(0, targetT - stepT);
      point = getPointAlongPath(originX, originY, destinationX, destinationY, targetT);
      snappedPoint = getSnappedPoint(tokenDocument, point.x, point.y);
      backstepIterations += 1;
    }
  }
  return snappedPoint;
}

async function applyResolvedMovement(tokenDocument, updateData, x, y) {
  const resolvedUpdate = { ...updateData, x, y };
  await tokenDocument.update(resolvedUpdate, { [INTERNAL_MOVE_OPTION]: true });
}

function syncSqueezeConditions(actor, { squeezing, squeezed }) {
  if (!actor?.setCondition) return;
  (async () => {
    const hasSqueezing = actor.statuses?.has?.("squeezing") ?? false;
    const hasSqueezed = actor.statuses?.has?.("squeezed") ?? false;
    if (hasSqueezing !== squeezing) await actor.setCondition("squeezing", squeezing);
    if (hasSqueezed !== squeezed) await actor.setCondition("squeezed", squeezed);
  })().catch((err) => {
    console.error(`${MODULE.ID} | Failed to sync squeezing conditions`, err);
  });
}

export function handleSqueezingPreTokenUpdate(tokenDocument, updateData, options = {}, _userId) {
  if (!game.settings.get(MODULE.ID, "automateSqueezing")) {
    return;
  }
  if (options?.[INTERNAL_MOVE_OPTION]) {
    return true;
  }
  if (!hasMovementUpdate(updateData)) {
    return;
  }

  const originX = tokenDocument.x;
  const originY = tokenDocument.y;
  const destinationX = Number.isFinite(updateData.x) ? updateData.x : originX;
  const destinationY = Number.isFinite(updateData.y) ? updateData.y : originY;
  if (originX === destinationX && originY === destinationY) {
    return;
  }

  const squeezingGeometry = getSqueezingGeometry(tokenDocument);
  const {
    tokenOverallWidthPx,
    tokenHeadWidthPx,
    halfOverallWidthPx,
    escapeArtistDc
  } = squeezingGeometry;
  const passageMetrics = computePassageWidthMetrics(
    tokenDocument,
    originX,
    originY,
    destinationX,
    destinationY,
    tokenOverallWidthPx
  );
  const measuredPassageWidth = passageMetrics?.narrowestWidth ?? null;
  const endPassageWidth = passageMetrics?.endWidth ?? null;
  const endsInNarrowArea = Number.isFinite(endPassageWidth) && endPassageWidth < tokenOverallWidthPx;
  const squeezedExitHandling = game.settings.get(MODULE.ID, "squeezedExitHandling");
  const pathAnalysis = computePathStateAnalysis(
    tokenDocument,
    originX,
    originY,
    destinationX,
    destinationY,
    tokenOverallWidthPx,
    halfOverallWidthPx,
    tokenHeadWidthPx
  );
  const needsEscapeArtistRoll = pathAnalysis.firstSqueezedT !== null;
  const entersHeadBlockedArea = pathAnalysis.firstHeadBlockedT !== null && !pathAnalysis.startsInHeadBlocked;

  const fullReachable = isPathReachable(tokenDocument, originX, originY, destinationX, destinationY, 1);
  const halfReachable = isPathReachable(tokenDocument, originX, originY, destinationX, destinationY, HALF_SCALE);
  const quarterReachable = isPathReachable(tokenDocument, originX, originY, destinationX, destinationY, QUARTER_SCALE);
  const movementActor = tokenDocument.actor ?? tokenDocument.object?.actor;

  let shouldApplySystemSqueezing = false;
  let shouldApplyNasSqueezed = false;
  if (Number.isFinite(measuredPassageWidth)) {
    if (measuredPassageWidth < tokenOverallWidthPx && measuredPassageWidth >= halfOverallWidthPx) {
      shouldApplySystemSqueezing = true;
    } else if (measuredPassageWidth < halfOverallWidthPx) {
      shouldApplyNasSqueezed = true;
    }
  }

  let expectedResult;
  if (!Number.isFinite(measuredPassageWidth)) {
    // Fallback if no enclosing walls were detected around the sampled path.
    expectedResult = "normal movement (no squeeze condition)";
  } else if (shouldApplySystemSqueezing) {
    expectedResult = "apply system squeezing";
  } else if (shouldApplyNasSqueezed) {
    expectedResult = "apply NAS squeezed";
  } else {
    expectedResult = "normal movement (no squeeze condition)";
  }

  if (entersHeadBlockedArea) {
    if (!pathAnalysis.startsInSqueezed) {
      shouldApplySystemSqueezing = false;
      shouldApplyNasSqueezed = false;
    }
    expectedResult = "blocked (space smaller than configured head width)";
  }

  if (squeezedExitHandling === EXIT_HANDLING_IGNORE && !endsInNarrowArea) {
    shouldApplySystemSqueezing = false;
    shouldApplyNasSqueezed = false;
    expectedResult = "normal movement (destination not narrow and ignore setting)";
  }

  if (fullReachable === null || halfReachable === null || quarterReachable === null) {
    expectedResult = `${expectedResult} | movement probe unavailable`;
  } else if (!quarterReachable) {
    shouldApplySystemSqueezing = false;
    shouldApplyNasSqueezed = false;
    expectedResult = "blocked (too narrow)";
  }

  if (entersHeadBlockedArea) {
    syncSqueezeConditions(movementActor, {
      squeezing: shouldApplySystemSqueezing,
      squeezed: shouldApplyNasSqueezed
    });
    if (options) options.cancelled = true;
    const headFailPoint = getEscapeFailurePoint(
      tokenDocument,
      originX,
      originY,
      destinationX,
      destinationY,
      pathAnalysis.firstHeadBlockedT,
      pathAnalysis.sampleSteps,
      ESCAPE_FAIL_STOP_BEFORE
    );
    void applyResolvedMovement(tokenDocument, updateData, headFailPoint.x, headFailPoint.y).catch((err) => {
      console.error(`${MODULE.ID} | Head-width movement handling failed`, err);
    });
    return false;
  }

  if (needsEscapeArtistRoll && !pathAnalysis.startsInSqueezed) {
    const actor = tokenDocument.actor ?? tokenDocument.object?.actor;
    if (!actor) return true;

    const failureHandling = game.settings.get(MODULE.ID, "squeezedEscapeFailureHandling");
    if (options) options.cancelled = true;

    (async () => {
      const rollResult = await actor.rollSkill(ESCAPE_ARTIST_SKILL, {
        dc: escapeArtistDc,
        reason: "squeezed"
      });
      const total =
        rollResult?.rolls?.[0]?.total
        ?? rollResult?.total
        ?? 0;

      if (total >= escapeArtistDc) {
        syncSqueezeConditions(actor, {
          squeezing: shouldApplySystemSqueezing,
          squeezed: shouldApplyNasSqueezed
        });
        await applyResolvedMovement(tokenDocument, updateData, destinationX, destinationY);
        return;
      }

      const failConditionState = failureHandling === ESCAPE_FAIL_ENTER_FIRST
        ? { squeezing: false, squeezed: true }
        : { squeezing: false, squeezed: false };
      syncSqueezeConditions(actor, failConditionState);
      const failPoint = getEscapeFailurePoint(
        tokenDocument,
        originX,
        originY,
        destinationX,
        destinationY,
        pathAnalysis.firstSqueezedT,
        pathAnalysis.sampleSteps,
        failureHandling
      );
      await applyResolvedMovement(tokenDocument, updateData, failPoint.x, failPoint.y);
    })().catch((err) => {
      console.error(`${MODULE.ID} | Escape Artist movement handling failed`, err);
    });

    return false;
  }

  syncSqueezeConditions(movementActor, {
    squeezing: shouldApplySystemSqueezing,
    squeezed: shouldApplyNasSqueezed
  });
}

// Temporary export alias for existing hook wiring compatibility.
export const handleSqueezedPreTokenUpdate = handleSqueezingPreTokenUpdate;
