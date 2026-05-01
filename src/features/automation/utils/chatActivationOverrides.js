import { MODULE } from "../../../common/module.js";
import { elementFromHtmlLike } from "../../../common/foundryCompat.js";

function getCastingTimeLabel() {
  return (game.i18n?.localize?.("PF1.CastingTime") ?? "Casting Time").toLowerCase();
}

function getActivationLabel(activation) {
  if (!activation) return "";
  const parser = pf1?.utils?.chat?.parseActivationLabel;
  if (typeof parser === "function") {
    return parser(activation) ?? "";
  }
  return "";
}

function readExistingActivationText(strong) {
  if (!strong) return "";
  const parts = [];
  let node = strong.nextSibling;
  while (node) {
    if (node.nodeType === Node.ELEMENT_NODE && node.nodeName === "BR") break;
    if (node.nodeType === Node.TEXT_NODE) {
      parts.push(node.textContent ?? "");
    } else if (node.nodeType === Node.ELEMENT_NODE) {
      parts.push(node.textContent ?? "");
    }
    node = node.nextSibling;
  }
  return parts.join("").replace(/\s+/g, " ").trim();
}

function replaceActivationText(strong, nextLabel) {
  if (!strong) return;
  let node = strong.nextSibling;
  while (node) {
    const next = node.nextSibling;
    if (node.nodeType === Node.ELEMENT_NODE && node.nodeName === "BR") break;
    node.remove();
    node = next;
  }
  strong.after(document.createTextNode(` ${nextLabel}`));
}

export function applyChatActivationOverrides(message, html) {
  const override = message?.flags?.[MODULE.ID]?.actionOverrides?.activation;
  if (!override) return;

  const activation = override.activation ?? message?.actionSource?.activation;
  const label = getActivationLabel(activation);

  const root = elementFromHtmlLike(html);
  if (!root) return;

  const castingTimeLabel = getCastingTimeLabel();
  const strongs = root.querySelectorAll("strong");
  for (const strong of strongs) {
    const text = (strong.textContent ?? "").trim().toLowerCase();
    if (text !== castingTimeLabel && text !== "casting time") continue;
    const existingText = readExistingActivationText(strong);
    const baseText = label || existingText;
    if (!baseText) return;

    let nextLabel = baseText;
    if (override.extraFullRound) {
      const fullLabel = getActivationLabel({ type: "full", cost: 1 }) || "full-round";
      nextLabel = `${baseText} + ${fullLabel}`;
    }

    replaceActivationText(strong, nextLabel);
    break;
  }
}
