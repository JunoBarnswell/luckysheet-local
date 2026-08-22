import { getFocusedContext, getFocusedId, focusUnit, requireFocusedContext, getUnit } from "./registry";

let focusHook = function (id) {
    return focusUnit(id);
};

export function setFocusHook(fn) {
    if (typeof fn === "function") {
        focusHook = fn;
    }
}

function currentPrefix() {
    const ctx = getFocusedContext();
    return (ctx && ctx.domPrefix) || "";
}

export function getLsPrefix(ctx) {
    if (ctx && ctx.domPrefix != null) {
        return ctx.domPrefix;
    }
    return currentPrefix();
}

export function isWidgetId(id) {
    if (!id) {
        return false;
    }
    return /^(luckysheet|testdpidiv|cellDatePickerBtn)/.test(id);
}

export function lsId(name) {
    if (!name) {
        return name;
    }
    const prefix = currentPrefix();
    if (!prefix) {
        return name;
    }
    if (name.indexOf(prefix) === 0) {
        return name;
    }
    return prefix + name;
}

export function $ls(selOrName) {
    const $ = typeof window !== "undefined" ? window.jQuery : null;
    if (!$) {
        return null;
    }
    if (!selOrName) {
        return $([]);
    }
    const ctx = requireFocusedContext();
    if (selOrName.charAt(0) === ".") {
        return $("#" + ctx.container).find(selOrName);
    }
    const id = selOrName.charAt(0) === "#" ? selOrName.slice(1) : selOrName;
    return $("#" + lsId(id));
}

export function prefixHtmlIds(html, prefix) {
    const actualPrefix = prefix == null ? currentPrefix() : prefix;
    if (!html || !actualPrefix) {
        return html;
    }
    return String(html).replace(/\b(id|for|aria-controls)\s*=\s*(["'])([^"']*)\2/g, function (match, attr, quote, id) {
        if (!id || !isWidgetId(id) || id.indexOf(actualPrefix) === 0) {
            return match;
        }
        return attr + "=" + quote + actualPrefix + id + quote;
    });
}

export function scopeSelector(selector) {
    if (typeof selector !== "string") {
        return selector;
    }
    const prefix = currentPrefix();
    if (!prefix) {
        return selector;
    }
    return selector.replace(/#([A-Za-z_][A-Za-z0-9_-]*)/g, function (match, id) {
        if (!isWidgetId(id) || id.indexOf(prefix) === 0) {
            return match;
        }
        return "#" + prefix + id;
    });
}

export function markInstanceNodes($nodes, instanceId) {
    if (!$nodes || !$nodes.length || !instanceId) {
        return $nodes;
    }
    $nodes.each(function () {
        if (this && this.nodeType === 1) {
            this.setAttribute("data-ls-instance", instanceId);
        }
    });
    $nodes.find("*").attr("data-ls-instance", instanceId);
    return $nodes;
}

export function appendInstancePortal(node) {
    const context = requireFocusedContext();
    const element = node && node.jquery ? node : $(node);
    markInstanceNodes(element, context.instanceId);
    $("body").append(element);
    element.each(function () { context.portals.add(this); });
    return element;
}

export function bindContainerFocus(instanceId, container) {
    if (typeof document === "undefined" || !container) {
        return;
    }
    const el = document.getElementById(container);
    if (!el) {
        return;
    }
    el.setAttribute("data-ls-host", instanceId);
    if (el.__lsFocusBound) {
        return;
    }
    el.__lsFocusBound = true;
    const onPointerDown = function () {
        const id = el.getAttribute("data-ls-host");
        if (!id) {
            return;
        }
        if (getFocusedId() !== id) {
            focusHook(id);
        }
    };
    el.addEventListener("pointerdown", onPointerDown, true);
    const context = getUnit(instanceId);
    if (context) {
        context.disposers.push(function () {
            el.removeEventListener("pointerdown", onPointerDown, true);
            delete el.__lsFocusBound;
        });
    }
}
