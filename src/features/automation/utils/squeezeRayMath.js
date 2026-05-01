const HALF_SCALE = 0.5;
const QUARTER_SCALE = 0.25;
const SAMPLE_T_VALUES = [0, 0.5, 1];
const MIN_PATH_SAMPLES = 4;
const MAX_PATH_SAMPLES = 80;
const PROBE_RAY_COUNT = 4;

function getWallCoords(wall) {
  const coords = wall?.document?.c ?? wall?.coords;
  if (!Array.isArray(coords) || coords.length !== 4) return null;
  return { x0: coords[0], y0: coords[1], x1: coords[2], y1: coords[3] };
}

function isMovementBlockingWall(wall) {
  const wallMoveType = wall?.document?.move;
  const none = CONST?.WALL_MOVEMENT_TYPES?.NONE;
  if (none === undefined) return true;
  return wallMoveType !== none;
}

export function getTokenCenter(tokenDocument, x, y) {
  return tokenDocument.getCenterPoint({
    x,
    y,
    elevation: tokenDocument.elevation ?? 0,
    width: tokenDocument.width,
    height: tokenDocument.height,
    shape: tokenDocument.shape
  });
}

export function buildWaypoint(tokenDocument, x, y, scale) {
  const width = Math.max(QUARTER_SCALE, (tokenDocument.width ?? 1) * scale);
  const height = Math.max(QUARTER_SCALE, (tokenDocument.height ?? 1) * scale);
  return {
    x,
    y,
    elevation: tokenDocument.elevation ?? 0,
    width,
    height,
    shape: tokenDocument.shape,
    action: tokenDocument.movementAction ?? "walk",
    snapped: true,
    explicit: true,
    checkpoint: false
  };
}

export function isPathReachable(tokenDocument, originX, originY, destinationX, destinationY, scale) {
  const tokenObject = tokenDocument.object;
  if (!tokenObject?.constrainMovementPath) return null;

  const waypoints = [
    buildWaypoint(tokenDocument, originX, originY, scale),
    buildWaypoint(tokenDocument, destinationX, destinationY, scale)
  ];
  const [, constrained] = tokenObject.constrainMovementPath(waypoints, {
    preview: true,
    ignoreWalls: false,
    ignoreCost: true,
    history: false
  });
  return !constrained;
}

export function nearestWallIntersectionDistance(start, direction, maxDistance) {
  const walls = canvas?.walls?.placeables ?? [];
  if (!walls.length) return null;

  const rayEnd = {
    x: start.x + (direction.x * maxDistance),
    y: start.y + (direction.y * maxDistance)
  };

  let nearest = null;
  for (const wall of walls) {
    if (!isMovementBlockingWall(wall)) continue;
    const coords = getWallCoords(wall);
    if (!coords) continue;

    const hit = foundry.utils.lineSegmentIntersection(
      start,
      rayEnd,
      { x: coords.x0, y: coords.y0 },
      { x: coords.x1, y: coords.y1 }
    );
    if (!hit) continue;

    const dx = hit.x - start.x;
    const dy = hit.y - start.y;
    const distance = Math.hypot(dx, dy);
    if (distance <= 0.001) continue;
    if (nearest === null || distance < nearest) nearest = distance;
  }

  return nearest;
}

export function getProbeDirectionPairs(originX, originY, destinationX, destinationY) {
  const movementDx = destinationX - originX;
  const movementDy = destinationY - originY;
  const movementLength = Math.hypot(movementDx, movementDy);
  if (movementLength <= 0.001) return null;

  const baseAngle = Math.atan2(movementDy, movementDx);
  const pairCount = Math.max(1, Math.floor(PROBE_RAY_COUNT / 2));
  const pairs = [];
  for (let i = 0; i < pairCount; i += 1) {
    const angle = baseAngle + ((Math.PI / pairCount) * i);
    const first = { x: Math.cos(angle), y: Math.sin(angle) };
    const second = { x: -first.x, y: -first.y };
    pairs.push({ first, second });
  }
  return pairs;
}

export function measurePassageWidthAtCenter(sampleCenter, firstDirection, secondDirection, probeDistance) {
  const firstDistance = nearestWallIntersectionDistance(sampleCenter, firstDirection, probeDistance);
  const secondDistance = nearestWallIntersectionDistance(sampleCenter, secondDirection, probeDistance);
  if (!Number.isFinite(firstDistance) || !Number.isFinite(secondDistance)) return null;
  return firstDistance + secondDistance;
}

export function measurePassageWidthAtCenterWithDirectionPairs(sampleCenter, directionPairs, probeDistance) {
  if (!Array.isArray(directionPairs) || directionPairs.length === 0) return null;
  let bestWidth = null;
  for (const pair of directionPairs) {
    const width = measurePassageWidthAtCenter(sampleCenter, pair.first, pair.second, probeDistance);
    if (!Number.isFinite(width)) continue;
    if (bestWidth === null || width < bestWidth) bestWidth = width;
  }
  return bestWidth;
}

export function computePassageWidthMetrics(tokenDocument, originX, originY, destinationX, destinationY, tokenOverallWidthPx) {
  const directionPairs = getProbeDirectionPairs(originX, originY, destinationX, destinationY);
  if (!directionPairs) return null;

  const movementDx = destinationX - originX;
  const movementDy = destinationY - originY;

  const gridSize = canvas?.grid?.size ?? 100;
  const probeDistance = Math.max(tokenOverallWidthPx * 2, gridSize * 2);

  let narrowestWidth = Infinity;
  let startWidth = null;
  let endWidth = null;
  let sampleCount = 0;

  for (const t of SAMPLE_T_VALUES) {
    const sampleX = originX + (movementDx * t);
    const sampleY = originY + (movementDy * t);
    const sampleCenter = getTokenCenter(tokenDocument, sampleX, sampleY);
    if (!sampleCenter) continue;

    const width = measurePassageWidthAtCenterWithDirectionPairs(sampleCenter, directionPairs, probeDistance);
    if (!Number.isFinite(width)) continue;
    if (t === 0) startWidth = width;
    if (t === 1) endWidth = width;
    narrowestWidth = Math.min(narrowestWidth, width);
    sampleCount += 1;
  }

  if (sampleCount === 0) return null;
  return {
    narrowestWidth,
    startWidth,
    endWidth
  };
}

export function computeNarrowMovementCost(tokenDocument, originX, originY, destinationX, destinationY, tokenOverallWidthPx) {
  const movementDx = destinationX - originX;
  const movementDy = destinationY - originY;
  const totalPixels = Math.hypot(movementDx, movementDy);
  if (totalPixels <= 0.001) {
    return {
      rawDistance: 0,
      narrowDistance: 0,
      effectiveDistance: 0
    };
  }

  const directionPairs = getProbeDirectionPairs(originX, originY, destinationX, destinationY);
  if (!directionPairs) {
    return {
      rawDistance: 0,
      narrowDistance: 0,
      effectiveDistance: 0
    };
  }
  const gridSize = canvas?.grid?.size ?? 100;
  const gridDistance = canvas?.dimensions?.distance ?? canvas?.scene?.grid?.distance ?? 5;
  const probeDistance = Math.max(tokenOverallWidthPx * 2, gridSize * 2);
  const sampleSteps = Math.max(
    MIN_PATH_SAMPLES,
    Math.min(MAX_PATH_SAMPLES, Math.ceil(totalPixels / Math.max(1, gridSize / 2)))
  );
  const segmentPixels = totalPixels / sampleSteps;

  let narrowPixels = 0;
  for (let i = 0; i < sampleSteps; i += 1) {
    const t = (i + 0.5) / sampleSteps;
    const sampleX = originX + (movementDx * t);
    const sampleY = originY + (movementDy * t);
    const sampleCenter = getTokenCenter(tokenDocument, sampleX, sampleY);
    if (!sampleCenter) continue;

    const width = measurePassageWidthAtCenterWithDirectionPairs(sampleCenter, directionPairs, probeDistance);
    if (Number.isFinite(width) && width < tokenOverallWidthPx) {
      narrowPixels += segmentPixels;
    }
  }

  const pixelToDistance = gridDistance / gridSize;
  const rawDistance = totalPixels * pixelToDistance;
  const narrowDistance = narrowPixels * pixelToDistance;
  const effectiveDistance = rawDistance + narrowDistance;

  return {
    rawDistance,
    narrowDistance,
    effectiveDistance
  };
}

export function computePathStateAnalysis(
  tokenDocument,
  originX,
  originY,
  destinationX,
  destinationY,
  tokenOverallWidthPx,
  halfOverallWidthPx,
  tokenHeadWidthPx
) {
  const movementDx = destinationX - originX;
  const movementDy = destinationY - originY;
  const totalPixels = Math.hypot(movementDx, movementDy);
  if (totalPixels <= 0.001) {
    return {
      startsInSqueezed: false,
      endsInSqueezed: false,
      firstSqueezedT: null,
      startsInHeadBlocked: false,
      endsInHeadBlocked: false,
      firstHeadBlockedT: null,
      sampleSteps: 0
    };
  }

  const directionPairs = getProbeDirectionPairs(originX, originY, destinationX, destinationY);
  if (!directionPairs) {
    return {
      startsInSqueezed: false,
      endsInSqueezed: false,
      firstSqueezedT: null,
      startsInHeadBlocked: false,
      endsInHeadBlocked: false,
      firstHeadBlockedT: null,
      sampleSteps: 0
    };
  }
  const gridSize = canvas?.grid?.size ?? 100;
  const probeDistance = Math.max(tokenOverallWidthPx * 2, gridSize * 2);
  const sampleSteps = Math.max(
    MIN_PATH_SAMPLES,
    Math.min(MAX_PATH_SAMPLES, Math.ceil(totalPixels / Math.max(1, gridSize / 2)))
  );

  let firstSqueezedT = null;
  let startsInSqueezed = false;
  let endsInSqueezed = false;
  let firstHeadBlockedT = null;
  let startsInHeadBlocked = false;
  let endsInHeadBlocked = false;

  for (let i = 0; i <= sampleSteps; i += 1) {
    const t = i / sampleSteps;
    const sampleX = originX + (movementDx * t);
    const sampleY = originY + (movementDy * t);
    const sampleCenter = getTokenCenter(tokenDocument, sampleX, sampleY);
    if (!sampleCenter) continue;
    const width = measurePassageWidthAtCenterWithDirectionPairs(sampleCenter, directionPairs, probeDistance);
    const isSqueezed = Number.isFinite(width) && width < halfOverallWidthPx;
    const isHeadBlocked = Number.isFinite(width) && width < tokenHeadWidthPx;
    if (i === 0) startsInSqueezed = isSqueezed;
    if (i === sampleSteps) endsInSqueezed = isSqueezed;
    if (i === 0) startsInHeadBlocked = isHeadBlocked;
    if (i === sampleSteps) endsInHeadBlocked = isHeadBlocked;
    if (isSqueezed && firstSqueezedT === null) firstSqueezedT = t;
    if (isHeadBlocked && firstHeadBlockedT === null) firstHeadBlockedT = t;
  }

  return {
    startsInSqueezed,
    endsInSqueezed,
    firstSqueezedT,
    startsInHeadBlocked,
    endsInHeadBlocked,
    firstHeadBlockedT,
    sampleSteps
  };
}

export {
  HALF_SCALE,
  QUARTER_SCALE,
  SAMPLE_T_VALUES,
  MIN_PATH_SAMPLES,
  MAX_PATH_SAMPLES
};
