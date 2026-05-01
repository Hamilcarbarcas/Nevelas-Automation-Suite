export const METAMAGIC_DEFINITION = {
  key: "quickenSpell",
  name: "Quicken Spell",
  get prefix() { return globalThis.game?.i18n?.localize?.("NAS.metamagic.prefixes.quickenSpell") ?? "Quicken"; },
};

export const QUICKEN_SPELL_NAME = METAMAGIC_DEFINITION.name;

export function applyQuickenSpell(context) {
  if (!context?.activation) return false;

  context.activation.type = "swift";
  context.activation.cost = 1;
  if (context.activation.unchained) {
    context.activation.unchained.type = "swift";
    context.activation.unchained.cost = 1;
  }

  if (!context.metamagic) {
    context.metamagic = { applied: [], slotIncrease: 0 };
  }

  if (!context.metamagic.applied.includes(METAMAGIC_DEFINITION.name)) {
    context.metamagic.applied.push(METAMAGIC_DEFINITION.name);
    context.metamagic.slotIncrease += 4;
  }

  return true;
}
