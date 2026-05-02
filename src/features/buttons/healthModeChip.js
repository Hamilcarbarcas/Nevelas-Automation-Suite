import { isWoundsVigorAutomationEnabled } from '../automation/utils/woundsVigor.js';

const HEALTH_MODE_ATTR = "data-nas-health-mode";
const MODE_CHIP_ATTR = "data-nas-health-mode-chip";

export function addPerAttackHealthModeChip(root) {
    if (!(root instanceof HTMLElement)) return;
    if (!isWoundsVigorAutomationEnabled()) return;

    const attacks = root.querySelectorAll(".chat-attack");
    for (const attack of attacks) {
        const hasApplyButton = attack.querySelector('a.inline-action[data-action="applyDamage"],button[data-action="applyDamage"]');
        if (!hasApplyButton) continue;

        if (!attack.getAttribute(HEALTH_MODE_ATTR)) {
            attack.setAttribute(HEALTH_MODE_ATTR, "vigor");
        }

        if (attack.querySelector(`[${MODE_CHIP_ATTR}]`)) continue;

        const anchorInfo = resolveChipAnchor(attack);
        const anchor = anchorInfo?.anchor;
        if (!(anchor instanceof HTMLElement)) continue;

        const chip = document.createElement("a");
        chip.setAttribute(MODE_CHIP_ATTR, "1");
        chip.dataset.mode = "vigor";
        chip.dataset.tooltip = "NAS.damage.HealthModeVigor";
        chip.classList.add("inline-action");
        chip.style.cssText = "padding:0 4px;line-height:1.1;font-size:11px;border:1px solid rgba(120,120,120,.45);border-radius:9px;display:inline-flex;align-items:center;justify-content:center;min-width:20px;height:18px;vertical-align:middle;";
        chip.style.cursor = "pointer";
        chip.style.opacity = "0.92";
        chip.innerHTML = '<i class="fa-solid fa-heart-pulse" inert=""></i>';
        applyChipVisual(chip, "vigor");

        chip.addEventListener("click", (ev) => {
            ev.preventDefault();
            ev.stopPropagation();
            const mode = attack.getAttribute(HEALTH_MODE_ATTR) === "wounds" ? "vigor" : "wounds";
            attack.setAttribute(HEALTH_MODE_ATTR, mode);
            chip.dataset.mode = mode;
            chip.dataset.tooltip = mode === "wounds"
                ? "NAS.damage.HealthModeWounds"
                : "NAS.damage.HealthModeVigor";
            applyChipVisual(chip, mode);
            refreshTooltipLive(chip);
        });

        const wrap = document.createElement("span");
        wrap.appendChild(chip);
        mountChip(anchor, wrap, anchorInfo.mode);
    }
}

export function getHealthModeForElement(el) {
    if (!isWoundsVigorAutomationEnabled()) return "vigor";
    const attack = el?.closest?.(".chat-attack");
    return attack?.getAttribute?.(HEALTH_MODE_ATTR) === "wounds" ? "wounds" : "vigor";
}

function applyChipVisual(chip, mode) {
    if (!(chip instanceof HTMLElement)) return;
    if (mode === "wounds") {
        chip.innerHTML = '<i class="fa-solid fa-droplet" inert=""></i>';
        chip.style.color = "#ef1313";
        chip.style.borderColor = "rgba(239,107,107,.65)";
        chip.style.boxShadow = "0 0 0 1px rgba(239,107,107,.26) inset";
    } else {
        chip.innerHTML = '<i class="fa-solid fa-heart-pulse" inert=""></i>';
        chip.style.color = "#ef1313";
        chip.style.borderColor = "rgba(255,143,143,.62)";
        chip.style.boxShadow = "0 0 0 1px rgba(239,107,107,.26) inset";
    }
    const icon = chip.querySelector("i");
    if (icon instanceof HTMLElement) {
        icon.style.textShadow = "0 0 1px rgba(0,0,0,.95), 0 0 2px rgba(0,0,0,.6)";
    }
}

function refreshTooltipLive(el) {
    if (!(el instanceof HTMLElement)) return;
    if (!el.matches(":hover")) return;
    const tooltipMgr = game?.tooltip;
    if (!tooltipMgr) return;
    tooltipMgr.clearPending();
    tooltipMgr.deactivate();
    tooltipMgr.activate(el);
}

function resolveChipAnchor(attack) {
    const flavorAnchor =
        attack.querySelector("th.attack-flavor[colspan='4']") ??
        attack.querySelector("thead tr:first-child th.attack-flavor") ??
        attack.querySelector(".attack-header");
    if (flavorAnchor instanceof HTMLElement) {
        return { anchor: flavorAnchor, mode: "absolute" };
    }

    // Damage-only cards (no attack roll) use attack-damage headers.
    const damageAnchor =
        attack.querySelector("th.attack-damage[colspan='2']") ??
        attack.querySelector("thead tr:first-child th.attack-damage") ??
        attack.querySelector("th.attack-damage");
    if (damageAnchor instanceof HTMLElement) {
        return { anchor: damageAnchor, mode: "inline" };
    }

    return null;
}

function mountChip(anchor, wrap, mode) {
    if (!(anchor instanceof HTMLElement) || !(wrap instanceof HTMLElement)) return;

    if (mode === "absolute") {
        anchor.style.position = "relative";
        wrap.style.cssText = "position:absolute;right:6px;top:50%;transform:translateY(-50%);display:inline-flex;align-items:center;gap:4px;z-index:2;";
        anchor.appendChild(wrap);
        return;
    }

    // Damage-only cards: place the mode chip to the LEFT of the "Damage" label
    // so it is visually distinct from the apply/half/healing action buttons.
    wrap.style.cssText = "display:inline-flex;align-items:center;margin-right:6px;margin-left:2px;vertical-align:middle;";
    anchor.insertBefore(wrap, anchor.firstChild);
}
