import { MODULE } from '../../common/module.js';

export function handleCombatTrackerRender(app, html, data) {
  if (!game.settings.get(MODULE.ID, 'autoApplyFF')) return;
  const surpriseSelector = '[data-control="pf1ic-surprise-round"]';
  if (!data.combat) {
    if (typeof html.find === 'function') {
      html.find(surpriseSelector).remove();
    } else {
      html.querySelectorAll(surpriseSelector).forEach(btn => btn.remove());
    }
    return;
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
    const surpriseRoundLabel = game.i18n.localize('NAS.conditions.main.SurpriseRound');
    const surpriseSelector = '[data-control="pf1ic-surprise-round"]';

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
      const placeholderNav = html.find('nav.combat-controls.add-placeholder');
      const footerNav = html.find('nav.combat-controls[data-application-part="footer"]');
      const baseControls = placeholderNav.length ? placeholderNav : footerNav.length ? footerNav : html.find('.combat-controls').first();

      baseControls.css('flex-direction', 'column');

      html.find(surpriseSelector).remove();
      surpriseRoundButton = $(
        `<a class="combat-control" data-control="pf1ic-surprise-round" aria-label="${surpriseRoundLabel}" role="button">
          ${surpriseRoundLabel}
        </a>`
      );

      if (placeholderNav.length) {
        placeholderNav.append(surpriseRoundButton);
      } else {
        const beginCombatButton = footerNav.find('a[data-control="startCombat"], button[data-action="startCombat"]');
        if (beginCombatButton.length) {
          beginCombatButton.before(surpriseRoundButton);
        } else {
          baseControls.prepend(surpriseRoundButton);  
        }
      }

      surpriseRoundButton.off('.pf1ic').on('click.pf1ic', handleSurpriseClick);
    } else {
      const placeholderNav = html.querySelector('nav.combat-controls.add-placeholder');
      const footerNav = html.querySelector('nav.combat-controls[data-application-part="footer"]');
      const baseControls = placeholderNav ?? footerNav ?? html.querySelector('.combat-controls');

      if (baseControls) baseControls.style.flexDirection = 'column';

      html.querySelectorAll(surpriseSelector).forEach(btn => btn.remove());
      surpriseRoundButton = document.createElement('button');
      surpriseRoundButton.classList.add('combat-control');
      surpriseRoundButton.setAttribute('aria-label', surpriseRoundLabel);
      surpriseRoundButton.dataset.control = 'pf1ic-surprise-round';
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

      surpriseRoundButton.addEventListener('click', handleSurpriseClick);
    }

    if (typeof combatControls.on === 'function') {
      combatControls.on('click', 'a[data-control="startCombat"]', async () => {
        const isSurprise = data.combat?.getFlag(MODULE.ID, 'isSurprise') || false;
        await resetExemptFlags(data.combat);
        if (isSurprise) {
          await data.combat?.setFlag(MODULE.ID, 'isSurprise', false);
        }
      });
    } else {
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
    }
  }
}



