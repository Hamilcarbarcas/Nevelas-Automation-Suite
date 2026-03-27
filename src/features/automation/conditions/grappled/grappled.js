import { MODULE } from '../../../../common/module.js';
import { GRAPPLE_FORM_KEY } from '../../utils/attackDialogControls.js';

const GRAPPLE_FLAG_KEY = "grappleContext";

export function isGrappleSelected(action) {
  return action?.formData?.[GRAPPLE_FORM_KEY] ?? action?.shared?.formData?.[GRAPPLE_FORM_KEY];
}

function getPrimaryTarget(action) {
  const targets = action?.shared?.targets ?? [];
  return targets[0];
}

function getPrimaryAttackTotal(action) {
  const chatAttack = action?.shared?.chatAttacks?.[0];
  return chatAttack?.attack?.total ?? null;
}

function getTokenCmd(targetToken) {
  const cmd = targetToken?.actor?.system?.attributes?.cmd;
  if (!cmd) return null;
  const isFlatFooted = targetToken.actor.statuses?.has("flatFooted");
  return isFlatFooted ? cmd.flatFootedTotal : cmd.total;
}

async function storeGrappleFlag(targetActor, attackerActor, attackTotal, cmdValue) {
  if (!targetActor) return;
  await targetActor.setFlag(MODULE.ID, GRAPPLE_FLAG_KEY, {
    attacker: attackerActor?.uuid ?? null,
    attackTotal,
    cmd: cmdValue,
    timestamp: Date.now(),
  });
}

export async function handleGrappleResolution(action) {
  if (!isGrappleSelected(action)) return;
  if (!action?.action?.hasAttack) return;

  const targetToken = getPrimaryTarget(action);
  if (!targetToken) {
    ui.notifications?.info?.(game.i18n.localize('NAS.conditions.main.GrappleNoTargets'));
    return;
  }

  const attackTotal = getPrimaryAttackTotal(action);
  if (attackTotal === null || attackTotal === undefined) return;

  const cmdValue = getTokenCmd(targetToken);
  if (cmdValue === null || cmdValue === undefined) return;

  if (attackTotal > cmdValue) {
    const targetActor = targetToken.actor;
    const attackerActor = action.token?.actor ?? action.actor;
    await targetActor?.setCondition("grappled", true);
    await attackerActor?.setCondition("grappling", true);
    await storeGrappleFlag(targetActor, attackerActor, attackTotal, cmdValue);
  }
}

function getGrappleContext(actor) {
  return actor?.getFlag(MODULE.ID, GRAPPLE_FLAG_KEY);
}

export function buildGrappleConcentrationData(actor, action) {
  const spellLevel = action?.shared?.rollData?.sl ?? 0;
  const context = getGrappleContext(actor) || {};
  const grappleTotal = context.attackTotal ?? 0;
  const dc = 10 + spellLevel + grappleTotal;
  return { spellLevel, grappleTotal, dc };
}

export function handleGrappledActionUse(action) {
  const held = action.formData?.held;
  const token = action.token;
  const actor = token?.actor;

  const grappledHandling = game.settings.get(MODULE.ID, 'grappledHandling');
  if (grappledHandling && actor?.statuses.has("grappled") && held === "2h") {
    if (grappledHandling === "disabled") return;
    if (grappledHandling === "strict") {
      action.shared.reject = true;
      ui.notifications.info(game.i18n.format('NAS.conditions.main.GrappledTwoHands', { name: token.name }));
    } else if (grappledHandling === "lenient") {
      ui.notifications.info(game.i18n.format('NAS.conditions.main.GrappledLenient', { name: token.name }));
    }
  }
}

export async function handleGrappledConcentration(action) {
  const handling = game.settings.get(MODULE.ID, 'entangledGrappledHandling');
  if (handling !== 'grappled' && handling !== 'both') return;

  const itemSource = action.item;
  const token = action.token;
  const actor = token?.actor;
  if (!actor) return;
  if (itemSource.type !== "spell") return;

  const notifyConcentrationDc = (data) => {
    const msg = game.i18n.format('NAS.conditions.main.GrappleConcentrationDC', {
      dc: data.dc,
      spellLevel: data.spellLevel,
      grappleTotal: data.grappleTotal
    });
    ui.notifications.info(msg);
  };

  const rollWithContext = async (spellbook, reason) => {
    const data = buildGrappleConcentrationData(actor, action);
    const options = { skipDialog, dc: data.dc, reason };
    notifyConcentrationDc(data);
    await actor.rollConcentration(spellbook, options);
  };

  const skipDialog = game.user.isGM;

  if (actor.statuses.has("grappled") && itemSource.system.components?.somatic) {
    await rollWithContext(itemSource.system.spellbook, "grappled");
  }
}



