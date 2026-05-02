export function chatMessageStyle(styleName = "OTHER") {
  const styles = globalThis.CONST?.CHAT_MESSAGE_STYLES;
  if (styles?.[styleName] != null) return { style: styles[styleName] };

  const types = globalThis.CONST?.CHAT_MESSAGE_TYPES;
  if (types?.[styleName] != null) return { type: types[styleName] };

  return {};
}

export function isFoundryV13Plus() {
  const generation = Number(globalThis.game?.release?.generation);
  if (Number.isFinite(generation)) return generation >= 13;

  const version = String(globalThis.game?.version ?? "");
  return version ? foundry.utils.isNewerVersion(version, "12.999") : false;
}

export function chatRenderHookName() {
  if (globalThis.game?.release?.generation == null && !globalThis.game?.version) return null;
  return isFoundryV13Plus() ? "renderChatMessageHTML" : "renderChatMessage";
}

export function onRenderChatMessageCompat(handler) {
  const hookName = chatRenderHookName();
  if (!hookName) {
    return Hooks.once("init", () => onRenderChatMessageCompat(handler));
  }

  return Hooks.on(hookName, handler);
}

export function actorDocumentConstructor() {
  return globalThis.Actor ?? globalThis.foundry?.documents?.Actor ?? null;
}

export function isActorDocument(value) {
  if (!value) return false;

  const ActorCtor = actorDocumentConstructor();
  if (ActorCtor && value instanceof ActorCtor) return true;

  return value.documentName === "Actor" || value.constructor?.name === "ActorPF";
}

export function elementFromHtmlLike(html) {
  const HTMLElementCtor = globalThis.HTMLElement;

  if (HTMLElementCtor && html instanceof HTMLElementCtor) return html;
  if (Array.isArray(html)) return elementFromHtmlLike(html[0]);
  if (html?.jquery) return elementFromHtmlLike(html[0]);
  if (HTMLElementCtor && html?.[0] instanceof HTMLElementCtor) return html[0];
  if (typeof html?.querySelector === "function") return html;
  return null;
}

export function htmlElementFromRenderArg(html) {
  return elementFromHtmlLike(html);
}

export function jqueryFromHtmlLike(html) {
  if (html?.jquery) return html;
  const element = elementFromHtmlLike(html);
  if (!element || typeof globalThis.jQuery !== "function") return null;
  return globalThis.jQuery(element);
}

export function queryHtml(html, selector) {
  const root = elementFromHtmlLike(html);
  return root?.querySelector?.(selector) ?? null;
}

export function queryHtmlAll(html, selector) {
  const root = elementFromHtmlLike(html);
  return root?.querySelectorAll ? Array.from(root.querySelectorAll(selector)) : [];
}

export function closestHtml(html, selector) {
  const element = elementFromHtmlLike(html);
  return element?.closest?.(selector) ?? null;
}

export function setElementVisible(html, visible) {
  const element = elementFromHtmlLike(html);
  if (!element) return;
  element.style.display = visible ? "" : "none";
}

export function checkboxChecked(html) {
  const element = elementFromHtmlLike(html);
  return element?.checked === true;
}

export function insertNasSettingsSectionsContainer(tabEl, container) {
  if (isFoundryV13Plus()) {
    tabEl.prepend(container);
    return;
  }

  const titleEl = tabEl.querySelector(":scope > h2.border");
  if (titleEl) titleEl.insertAdjacentElement("afterend", container);
  else tabEl.prepend(container);
}
