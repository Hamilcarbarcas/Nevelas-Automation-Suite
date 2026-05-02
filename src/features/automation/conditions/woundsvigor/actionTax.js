import { MODULE } from '../../../../common/module.js';
import { isWoundsVigorActive, isWvNoWoundsActor, isWvWounded } from '../../utils/woundsVigor.js';

const _processedActionUses = new WeakSet();

const STRICT_ACTION_TYPES = new Set(['move', 'standard']);
const BROAD_ACTION_TYPES = new Set(['move', 'standard', 'full', 'attack', 'aoo', 'round']);

function extractRollTotal(rollResult) {
  if (!rollResult) return null;
  const direct = Number(rollResult.total);
  if (Number.isFinite(direct)) return direct;

  const first = Number(rollResult?.rolls?.[0]?.total);
  if (Number.isFinite(first)) return first;

  return null;
}

function isTriggeringActionType(activationType, mode) {
  const value = String(activationType ?? '').toLowerCase();
  if (!value) return false;
  if (mode === 'broad') return BROAD_ACTION_TYPES.has(value);
  return STRICT_ACTION_TYPES.has(value);
}

export async function handleWoundsVigorActionTax(action) {
  if (!action || _processedActionUses.has(action)) return;
  _processedActionUses.add(action);

  const actor = action?.token?.actor ?? action?.item?.actor ?? action?.actor;
  if (!actor?.isOwner) return;
  if (!isWoundsVigorActive(actor)) return;
  if (isWvNoWoundsActor(actor)) return;
  if (!isWvWounded(actor)) return;

  const mode = game.settings.get(MODULE.ID, 'woundsVigorActionTaxMode') || 'strict';
  const activationType = action?.action?.activation?.type;
  if (!isTriggeringActionType(activationType, mode)) return;

  await actor.applyDamage(1, {
    asWounds: true,
    dialog: false,
    nasWvActionTax: true
  });

  const result = await actor.rollAbilityTest('con', { skipDialog: true });
  const total = extractRollTotal(result);
  if (total === null) return;

  if (total < 10) {
    await actor.setCondition('unconscious', true);
    ui.notifications.warn(
      game.i18n.format('NAS.conditions.main.WoundsVigorConCheckFailed', { name: actor.name, total, dc: 10 })
    );
  }
}
