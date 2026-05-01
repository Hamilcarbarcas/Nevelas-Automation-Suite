import { MODULE } from '../../../../common/module.js';

export function handleSqueezedPreActionUse(action) {
  if (!game.settings.get(MODULE.ID, "automateSqueezing")) return;
  const actionType = action.action.activation?.type;
  const token = action.token;
  const actor = token?.actor;
  const squeezedHandling = game.settings.get(MODULE.ID, "squeezedHandling");
  if (squeezedHandling && actor?.statuses.has("squeezed")) {
    if (squeezedHandling === "disabled") return;
    if (squeezedHandling === "strict" && (actionType === "attack" || actionType === "aoo")) {
      action.shared.reject = true;
      ui.notifications.info(game.i18n.format("NAS.conditions.main.SqueezedStrict", { name: token.name }));
    } else if (squeezedHandling === "lenient") {
      ui.notifications.info(game.i18n.format("NAS.conditions.main.SqueezedLenient", { name: token.name }));
    }
  }
}
