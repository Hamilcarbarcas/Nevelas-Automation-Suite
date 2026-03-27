import { MODULE } from '../../../common/module.js';
import { customApplyDamage } from './applydamage.js';

const targetInfo = {};

function interceptCardData(wrapped, message, elementObject) {
    const actionName = elementObject.target?.dataset?.action;
    const button = elementObject.target;
    if ((actionName == "applyDamage" || actionName == "applyClusteredDamage") && button) {
        const chatMessage = button.closest('.chat-message');
        const chatAttack = button.closest('.chat-attack');

        if (chatMessage) {
            targetInfo.id = chatMessage.getAttribute('data-message-id');
        }
        targetInfo.buttonType = button.dataset.tooltip || button.innerText;

        if (chatAttack) {
            targetInfo.attackIndex = chatAttack.getAttribute('data-index');

            let damageElement = button.closest('th[data-damage-type]');

            if (!damageElement) {
                const buttonType = button.getAttribute('data-type'); 
                if (buttonType) {
                    damageElement = chatAttack.querySelector(`th[data-damage-type="${buttonType}"]`);
                }
            }

            if (damageElement) {
                targetInfo.isCritical = damageElement.getAttribute('data-damage-type') === 'critical';
            } else {
                targetInfo.isCritical = false;
            }
        }
    }
    return wrapped(message, elementObject);
}

export function registerLegacyDamageOverride() {
    function markDialogUse(wrapped, ...args) {
        const rv = wrapped(...args);
        rv._nasDamageDialog = true;
        return rv;
    }
    libWrapper.register(MODULE.ID, "pf1.applications.ApplyDamage.prototype._getTargetDamageOptions", markDialogUse, libWrapper.WRAPPER);
    libWrapper.register(MODULE.ID, 'pf1.utils.chat.onButton', interceptCardData, libWrapper.MIXED);

    libWrapper.register(MODULE.ID, "pf1.documents.actor.ActorPF.prototype.applyDamage", applyDamage, libWrapper.WRAPPER);
    function applyDamage(wrapped, value, config) {
        if (config._nasDamageDialog) return wrapped(value, config);
        const hasSelectedTokens = canvas.tokens.controlled.length > 0;
        const isClusteredShots = config.flags?.[MODULE.ID]?.clusteredShots != null;
        if (isClusteredShots && Object.keys(targetInfo).length === 0) {
            const clusteredShotsData = config.flags[MODULE.ID].clusteredShots;
            if (clusteredShotsData.message) {
                const message = clusteredShotsData.message;
                const buttonElement = clusteredShotsData.buttonElement;
                interceptCardData(function(){}, message, { target: buttonElement });
                if (Object.keys(targetInfo).length === 0) {
                    targetInfo.id = clusteredShotsData.messageId;
                    targetInfo.buttonType = "Clustered Shots";
                }
            }
        }
        const hasTargetInfo = Object.keys(targetInfo).length > 0;
        const isNormalDamage = !config.healing && value >= 0;
        if (hasSelectedTokens && hasTargetInfo && (isNormalDamage || isClusteredShots)) {
            customApplyDamage(wrapped, value, config, targetInfo);
        } else {
            return wrapped(value, config);
        }
    }
}
