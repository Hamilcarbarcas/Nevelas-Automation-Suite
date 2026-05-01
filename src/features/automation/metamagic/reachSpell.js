export const METAMAGIC_DEFINITION = {
  key: "reachSpell",
  name: "Reach Spell",
  get prefix() { return globalThis.game?.i18n?.localize?.("NAS.metamagic.prefixes.reachSpell") ?? "Reach"; },
};

const RANGE_ORDER = ["touch", "close", "medium", "long"];

function getNextRangeUnit(unit) {
  const index = RANGE_ORDER.indexOf(unit);
  if (index === -1 || index === RANGE_ORDER.length - 1) return null;
  return RANGE_ORDER[index + 1];
}

export function applyReachSpell(context, steps = 1) {
  const range = context?.range?.range;
  if (!range) return false;

  const units = (range.units ?? "").toString().toLowerCase();
  let currentUnit = units;
  let appliedSteps = 0;
  for (let i = 0; i < steps; i += 1) {
    const nextUnit = getNextRangeUnit(currentUnit);
    if (!nextUnit) break;
    currentUnit = nextUnit;
    appliedSteps += 1;
  }
  if (appliedSteps === 0) return false;

  range.units = currentUnit;
  context.range.hasRange = true;
  context.range.isRanged = true;
  context.range.touch = false;
  if (!range.minUnits || range.minUnits === "touch") {
    range.minUnits = currentUnit;
  }

  if (context.actionType === "msak") {
    context.actionType = "rsak";
  }

  if (!context.metamagic) {
    context.metamagic = { applied: [], slotIncrease: 0 };
  }

  if (!context.metamagic.applied.includes(METAMAGIC_DEFINITION.name)) {
    context.metamagic.applied.push(METAMAGIC_DEFINITION.name);
    context.metamagic.slotIncrease += appliedSteps;
  }

  return true;
}
