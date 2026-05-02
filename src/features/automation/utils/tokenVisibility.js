export function tokenCanSeeToken(sourceToken, targetToken) {
  if (!sourceToken || !targetToken) return false;
  if (!targetToken.center) return false;
  if (!canvas?.visibility?.testVisibility) return true;

  return canvas.visibility.testVisibility(targetToken.center, {
    object: targetToken,
    visionSource: sourceToken.vision
  });
}

export function tokenDistance(sourceToken, targetToken) {
  if (!sourceToken?.center || !targetToken?.center) return Infinity;
  return canvas?.grid?.measurePath?.([sourceToken.center, targetToken.center])?.distance ?? Infinity;
}
