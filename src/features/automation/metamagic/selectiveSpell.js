export const METAMAGIC_DEFINITION = {
  key: "selectiveSpell",
  name: "Selective Spell",
  get prefix() { return globalThis.game?.i18n?.localize?.("NAS.metamagic.prefixes.selectiveSpell") ?? "Selective"; },
};

export function applySelectiveSpell(context) {
  if (!context) return false;

  if (!context.metamagic) {
    context.metamagic = { applied: [], slotIncrease: 0 };
  }

  if (!context.metamagic.applied.includes(METAMAGIC_DEFINITION.name)) {
    context.metamagic.applied.push(METAMAGIC_DEFINITION.name);
    context.metamagic.slotIncrease += 1;
  }

  return true;
}
