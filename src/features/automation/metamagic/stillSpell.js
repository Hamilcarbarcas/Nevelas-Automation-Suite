export const METAMAGIC_DEFINITION = {
  key: "stillSpell",
  name: "Still Spell",
  get prefix() { return globalThis.game?.i18n?.localize?.("NAS.metamagic.prefixes.stillSpell") ?? "Still"; },
};

export function applyStillSpell(context) {
  if (!context?.components) return;

  if (context.components.somatic === true) {
    context.components.somatic = false;
  }

  if (!context.metamagic) {
    context.metamagic = { applied: [], slotIncrease: 0 };
  }

  if (!context.metamagic.applied.includes(METAMAGIC_DEFINITION.name)) {
    context.metamagic.applied.push(METAMAGIC_DEFINITION.name);
    context.metamagic.slotIncrease += 1;
  }
}
