import { MODULE } from '../common/module.js';
import { CONCEALMENT_CONDITION_ID, actorHasBlindFight, getConcealmentState } from '../features/automation/conditions/concealment/concealment.js';

export function registerConditionFootnoteWrapper(isGrappleSelected) {
  if (!game.modules.get("lib-wrapper")?.active) {
    console.warn(`${MODULE.ID} | libWrapper missing; grapple footnotes disabled.`);
    return;
  }

  libWrapper.register(
    MODULE.ID,
    "pf1.actionUse.ActionUse.prototype.addFootnotes",
    async function (wrapped, ...args) {
      await wrapped(...args);
      try {
        if (!this.action?.hasAttack) return;
        const actor = this.token?.actor ?? this.actor;

        if (isGrappleSelected(this) && actor?.statuses?.has?.("grappling")) {
          this.shared ??= {};
          this.shared.templateData ??= {};
          const footnotes = this.shared.templateData.footnotes ?? [];
          const text = game.i18n.localize('NAS.conditions.main.GrappleFootnote');
          if (!footnotes.some((entry) => entry?.text === text)) {
            footnotes.push({ text });
          }
          this.shared.templateData.footnotes = footnotes;
        }

        if (Array.isArray(this.shared?.targets) && this.shared.targets.length > 0) {
          let chosen = null;
          for (const target of this.shared.targets) {
            const actor = target?.actor;
            if (!actor) continue;
            if (!actor.statuses?.has?.(CONCEALMENT_CONDITION_ID)) continue;
            const state = getConcealmentState(actor) || {};
            const hasTotal = actor.items.some(i => i.type === "buff" && i.name === "Concealment (Total)" && i.system?.active);
            const hasNormal = actor.items.some(i => i.type === "buff" && i.name === "Concealment" && i.system?.active);
            if (!hasTotal && !hasNormal) continue; 
            const variant = state.variant || (hasTotal ? "total" : "normal");
            const targetData = { variant };
            chosen = (!chosen || chosen.variant === "normal" && variant === "total") ? targetData : chosen;
            if (chosen.variant === "total") break; 
          }

          if (chosen) {
            this.shared ??= {};
            this.shared.templateData ??= {};
            const footnotes = this.shared.templateData.footnotes ?? [];
            const hasBF = actorHasBlindFight(actor);
            const roll = hasBF ? "[[2d100kh]]" : "[[1d100]]";
            const threshold = chosen.variant === "total" ? 50 : 20;
            const variantLabel = chosen.variant === "total"
              ? game.i18n.localize('NAS.conditions.main.ConcealmentTotal')
              : game.i18n.localize('NAS.conditions.main.ConcealmentNormal');
            const text = game.i18n.format('NAS.conditions.main.ConcealmentFootnote', {
              variant: variantLabel,
              roll,
              threshold
            });
            if (!footnotes.some((entry) => entry?.text === text)) {
              footnotes.push({ text });
            }
            this.shared.templateData.footnotes = footnotes;
          }
        }
      } catch (err) {
        console.error(`${MODULE.ID} | Failed to append grapple footnote`, err);
      }
    },
    "WRAPPER"
  );
}

