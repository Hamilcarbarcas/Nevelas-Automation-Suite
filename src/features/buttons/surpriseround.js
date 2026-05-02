import { MODULE } from '../../common/module.js';
import { elementFromHtmlLike, isFoundryV13Plus } from '../../common/foundryCompat.js';

const SURPRISED_FLAG = "surprisedInSurpriseRound";
const SURPRISE_TOGGLE_SELECTOR = '[data-action="nas-toggle-surprise-eligible"]';

function localizeWithFallback(key, fallback) {
  const localized = game.i18n.localize(key);
  return localized === key ? fallback : localized;
}

function isCombatantMarkedSurprised(combatant) {
  const flagValue = combatant?.getFlag?.(MODULE.ID, SURPRISED_FLAG);
  return flagValue === undefined ? false : Boolean(flagValue);
}

async function setCombatantSurprised(combatant, value) {
  if (!combatant?.setFlag) return;
  await combatant.setFlag(MODULE.ID, SURPRISED_FLAG, Boolean(value));
}

function updateEligibilityButtonState(button, surprised) {
  if (!button) return;
  button.classList.toggle("active", surprised);
  button.setAttribute("aria-pressed", surprised ? "true" : "false");
}

export function handleCombatTrackerRender(app, html, data) {
  if (!game.settings.get(MODULE.ID, 'autoApplyFF')) return;
  html = elementFromHtmlLike(html);
  if (!html) return;
  const surpriseSelector = '[data-control="nas-surprise-round"]';
  const surprisedLabel = localizeWithFallback("NAS.combat.buttons.surprised.name", "Surprised");
  const surprisedTooltip = localizeWithFallback("NAS.combat.buttons.surprised.tooltip", "Surprised");
  const surpriseRoundLabel = localizeWithFallback("NAS.combat.buttons.surpriseRound.name", game.i18n.localize('NAS.conditions.main.SurpriseRound'));
  const surpriseRoundTooltip = localizeWithFallback("NAS.combat.buttons.surpriseRound.tooltip", surpriseRoundLabel);
  const useV13Controls = isFoundryV13Plus();
  if (!data.combat) {
    if (typeof html.find === 'function') {
      html.find(surpriseSelector).remove();
    } else {
      html.querySelectorAll(surpriseSelector).forEach(btn => btn.remove());
    }
    return;
  }

  if (game.user.isGM) {
    const canEditEligibility = data.combat?.current?.round === 0;
    const listSelector = 'li.combatant[data-combatant-id]';

    if (typeof html.find === "function") {
      html.find(`${listSelector} .combatant-controls`).each((_idx, el) => {
        const controls = el;
        const combatantId = controls.closest("li.combatant")?.dataset?.combatantId;
        if (!combatantId) return;
        const combatant = data.combat.combatants.get(combatantId);
        if (!combatant) return;

        $(controls).find(SURPRISE_TOGGLE_SELECTOR).remove();

        const surprised = isCombatantMarkedSurprised(combatant);
        const btn = $(
          `<a class="combatant-control" data-action="nas-toggle-surprise-eligible" data-tooltip="${surprisedTooltip}" aria-label="${surprisedLabel}" title="${surprisedTooltip}" role="button">
            <i class="fa-solid fa-person-falling-burst"></i>
          </a>`
        );
        btn.toggleClass("disabled", !canEditEligibility);
        btn.attr("aria-disabled", canEditEligibility ? "false" : "true");
        btn.on("click.nas", async (ev) => {
          ev.preventDefault();
          ev.stopPropagation();
          if (!canEditEligibility) return;
          await setCombatantSurprised(combatant, !isCombatantMarkedSurprised(combatant));
          app?.render?.();
        });

        const effects = controls.querySelector(".token-effects");
        if (effects?.parentNode) {
          effects.parentNode.insertBefore(btn[0], effects);
        } else {
          controls.appendChild(btn[0]);
        }
        updateEligibilityButtonState(btn[0], surprised);
      });
    } else {
      html.querySelectorAll(`${listSelector} .combatant-controls`).forEach((controls) => {
        const combatantId = controls.closest("li.combatant")?.dataset?.combatantId;
        if (!combatantId) return;
        const combatant = data.combat.combatants.get(combatantId);
        if (!combatant) return;

        controls.querySelectorAll(SURPRISE_TOGGLE_SELECTOR).forEach((b) => b.remove());

        const surprised = isCombatantMarkedSurprised(combatant);
        const btn = useV13Controls ? document.createElement("button") : document.createElement("a");
        if (useV13Controls) {
          btn.type = "button";
          btn.className = "inline-control combatant-control icon fa-solid fa-person-falling-burst";
          btn.disabled = !canEditEligibility;
        } else {
          const referenceControl = controls.querySelector("a.combatant-control, .combatant-control");
          btn.className = referenceControl?.className || "combatant-control";
          btn.setAttribute("role", "button");
          btn.classList.toggle("disabled", !canEditEligibility);
          btn.setAttribute("aria-disabled", canEditEligibility ? "false" : "true");
          btn.innerHTML = '<i class="fa-solid fa-person-falling-burst"></i>';
        }
        btn.dataset.action = "nas-toggle-surprise-eligible";
        btn.dataset.tooltip = surprisedTooltip;
        btn.setAttribute("aria-label", surprisedLabel);
        btn.title = surprisedTooltip;
        btn.addEventListener("click", async (ev) => {
          ev.preventDefault();
          ev.stopPropagation();
          if (!canEditEligibility) return;
          await setCombatantSurprised(combatant, !isCombatantMarkedSurprised(combatant));
          app?.render?.();
        });

        const effects = controls.querySelector(".token-effects");
        if (effects?.parentNode) {
          effects.parentNode.insertBefore(btn, effects);
        } else {
          controls.appendChild(btn);
        }
        updateEligibilityButtonState(btn, surprised);
      });
    }
  }
  let combatControls;
  if (typeof html.find === "function") {
    combatControls = html.find('nav.combat-controls[data-application-part="footer"]');
    if (!combatControls.length) combatControls = html.find('.combat-controls').first();
  } else {
    combatControls = html.querySelector('nav.combat-controls[data-application-part="footer"]') || html.querySelector('.combat-controls');
  }

  const isSurprise = data.combat?.getFlag(MODULE.ID, 'isSurprise') || false;
  const isRoundOne = data.combat?.current?.round === 1;
  
  if (isSurprise && isRoundOne) {
    let roundDisplay;
    const surpriseRoundText = game.i18n.localize('NAS.conditions.main.SurpriseRound');
    if (typeof html.find === "function") {
      roundDisplay = html.find('.encounter-title');
      if (roundDisplay.length) {
        roundDisplay.html(`<span style="color: red; font-weight: bold;">${surpriseRoundText}</span>`);
      }
    } else {
      roundDisplay = html.querySelector('.encounter-title');
      if (roundDisplay) {
        roundDisplay.innerHTML = `<span style="color: red; font-weight: bold;">${surpriseRoundText}</span>`;
      }
    }
  }

  if (data.combat?.current?.round === 0 && game.user.isGM) {
    const surpriseSelector = '[data-control="nas-surprise-round"]';

    const resetExemptFlags = async (combat) => {
      const selectedTokens = canvas.tokens.controlled.map(token => token.id); 

      const flagPromises = combat.turns.map(async turn => {
        const tokenId = turn.tokenId;
        const token = canvas.tokens.get(tokenId);

        if (token) {
          const isSelected = selectedTokens.includes(tokenId);
          return token.actor.setFlag(MODULE.ID, 'exemptFromSurprise', isSelected);
        }
      });

      await Promise.all(flagPromises);
    };

    const handleSurpriseClick = async () => {
      const isSurprise = data.combat?.getFlag(MODULE.ID, 'isSurprise') || false;

      await resetExemptFlags(data.combat);

      if (!isSurprise) {
        await data.combat?.setFlag(MODULE.ID, 'isSurprise', true);
      }
      data.combat?.startCombat();
    };

    let surpriseRoundButton;

    if (typeof html.find === 'function') {
      html.find(surpriseSelector).remove();
      surpriseRoundButton = $(
        `<a class="combat-control" data-control="nas-surprise-round" data-tooltip="${surpriseRoundTooltip}" aria-label="${surpriseRoundLabel}" role="button">
          ${surpriseRoundLabel}
        </a>`
      );

      const beginCombatButton = html.find('a[data-control="startCombat"], button[data-action="startCombat"]').first();
      if (beginCombatButton.length) {
        beginCombatButton.before(surpriseRoundButton);
      }

      surpriseRoundButton.off('.nas').on('click.nas', handleSurpriseClick);
    } else {
      html.querySelectorAll(surpriseSelector).forEach(btn => btn.remove());

      if (useV13Controls) {
        const placeholderNav = html.querySelector('nav.combat-controls.add-placeholder');
        const footerNav = html.querySelector('nav.combat-controls[data-application-part="footer"]');
        const baseControls = placeholderNav ?? footerNav ?? html.querySelector('.combat-controls');
        if (baseControls) baseControls.style.flexDirection = 'column';

        surpriseRoundButton = document.createElement('button');
        surpriseRoundButton.classList.add('combat-control');
        surpriseRoundButton.setAttribute('aria-label', surpriseRoundLabel);
        surpriseRoundButton.dataset.tooltip = surpriseRoundTooltip;
        surpriseRoundButton.dataset.control = 'nas-surprise-round';
        surpriseRoundButton.type = 'button';
        surpriseRoundButton.textContent = surpriseRoundLabel;

        if (placeholderNav) {
          placeholderNav.appendChild(surpriseRoundButton); 
        } else if (footerNav) {
          const beginCombatButton = footerNav.querySelector('button[data-action="startCombat"], a[data-control="startCombat"]');
          if (beginCombatButton && beginCombatButton.parentNode) {
            beginCombatButton.parentNode.insertBefore(surpriseRoundButton, beginCombatButton);
          } else {
            footerNav.insertBefore(surpriseRoundButton, footerNav.firstChild);
          }
        } else if (baseControls) {
          baseControls.insertBefore(surpriseRoundButton, baseControls.firstChild);
        }
      } else {
        const startCombatButton = html.querySelector('button[data-action="startCombat"], a[data-control="startCombat"]');
        surpriseRoundButton = document.createElement('a');
        surpriseRoundButton.className = startCombatButton?.className || 'combat-control';
        surpriseRoundButton.setAttribute('role', 'button');
        surpriseRoundButton.setAttribute('aria-label', surpriseRoundLabel);
        surpriseRoundButton.dataset.tooltip = surpriseRoundTooltip;
        surpriseRoundButton.dataset.control = 'nas-surprise-round';
        surpriseRoundButton.textContent = surpriseRoundLabel;

        if (startCombatButton?.parentNode) {
          if (startCombatButton.parentElement) {
            startCombatButton.parentElement.style.flexDirection = 'column';
          }
          startCombatButton.parentNode.insertBefore(surpriseRoundButton, startCombatButton);
        }
      }

      surpriseRoundButton.addEventListener('click', handleSurpriseClick);
    }

    if (combatControls && typeof combatControls.on === 'function') {
      combatControls.on('click', 'a[data-control="startCombat"]', async () => {
        const isSurprise = data.combat?.getFlag(MODULE.ID, 'isSurprise') || false;
        await resetExemptFlags(data.combat);
        if (isSurprise) {
          await data.combat?.setFlag(MODULE.ID, 'isSurprise', false);
        }
      });
    } else if (combatControls) {
      const startCombatButton = combatControls.querySelector('button[data-action="startCombat"]');
      if (startCombatButton) {
        startCombatButton.addEventListener('click', async () => {
          const isSurprise = data.combat?.getFlag(MODULE.ID, 'isSurprise') || false;
          await resetExemptFlags(data.combat);
          if (isSurprise) {
            await data.combat?.setFlag(MODULE.ID, 'isSurprise', false);
          }
        });
      }
    } else if (!useV13Controls) {
      const startCombatButton = html.querySelector('button[data-action="startCombat"], a[data-control="startCombat"]');
      startCombatButton?.addEventListener('click', async () => {
        const isSurprise = data.combat?.getFlag(MODULE.ID, 'isSurprise') || false;
        await resetExemptFlags(data.combat);
        if (isSurprise) {
          await data.combat?.setFlag(MODULE.ID, 'isSurprise', false);
        }
      });
    }
  }
}
