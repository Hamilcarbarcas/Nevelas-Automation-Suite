import { MODULE } from '../../../../common/module.js';
import { htmlElementFromRenderArg, onRenderChatMessageCompat } from '../../../../common/foundryCompat.js';

let hookRegistered = false;

export function registerConfusionChatMessageHook() {
  if (hookRegistered) return;
  hookRegistered = true;

  onRenderChatMessageCompat((message, html, data) => {
    if (!game.settings.get(MODULE.ID, 'handleConfused')) return;

    const root = htmlElementFromRenderArg(html);
    if (!root) return;

    const tokenImgs = root.querySelectorAll('.NAS-token img');
    tokenImgs.forEach(img => {
      img.addEventListener('click', async ev => {
        const icToken = ev.currentTarget.closest('.NAS-token');
        const tokenUuid = icToken?.dataset.uuid;
        const tokenDocument = await fromUuid(tokenUuid);
        const token = canvas.tokens.get(tokenDocument.id);
        if (token) {
          token.control({releaseOthers: true});
          canvas.animatePan({x: token.center?.x, y: token.center?.y, duration: 1000});
        }
      });
      img.addEventListener('mouseenter', async ev => {
        const icToken = ev.currentTarget.closest('.NAS-token');
        const tokenUuid = icToken?.dataset.uuid;
        const tokenDocument = await fromUuid(tokenUuid);
        const token = canvas.tokens.get(tokenDocument.id);
        if (token) token._onHoverIn(ev);
      });
      img.addEventListener('mouseleave', async ev => {
        const icToken = ev.currentTarget.closest('.NAS-token');
        const tokenUuid = icToken?.dataset.uuid;
        const tokenDocument = await fromUuid(tokenUuid);
        const token = canvas.tokens.get(tokenDocument.id);
        if (token) token._onHoverOut(ev);
      });
    });
  });
}
