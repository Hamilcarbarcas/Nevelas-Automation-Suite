export const METAMAGIC_DEFINITION = {
  key: "persistentSpell",
  name: "Persistent Spell",
  get prefix() { return globalThis.game?.i18n?.localize?.("NAS.metamagic.prefixes.persistentSpell") ?? "Persistent"; },
};

const PERSISTENT_FOOTNOTE =
  "Persistent Spell: on a successful save, the creature must reroll the save; failing the second save applies full effects.";

export function applyPersistentSpell(context) {
  if (!context?.save?.type) return false;

  if (!context.metamagic) {
    context.metamagic = { applied: [], slotIncrease: 0 };
  }

  if (!context.notes) {
    context.notes = {};
  }
  if (!Array.isArray(context.notes.footer)) {
    context.notes.footer = [];
  }

  if (!context.notes.footer.includes(PERSISTENT_FOOTNOTE)) {
    context.notes.footer.push(PERSISTENT_FOOTNOTE);
  }

  if (!context.metamagic.applied.includes(METAMAGIC_DEFINITION.name)) {
    context.metamagic.applied.push(METAMAGIC_DEFINITION.name);
    context.metamagic.slotIncrease += 2;
  }

  context.metamagic.persistent = true;
  return true;
}
