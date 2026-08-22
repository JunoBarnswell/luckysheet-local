import { getFocusedContext, getFocusedId, focusUnit } from "./registry";

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
    return /^(luckysheet[-_]|luckysheetTableContent|luckysheetcoltable|luckysheetrowHeader)/.test(id);
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
    if (selOrName.charAt(0) === ".") {
        const ctx = getFocusedContext();
        if (ctx && ctx.container) {
            return $("#" + ctx.container).find(selOrName);
        }
        return $(selOrName);
    }
    const id = selOrName.charAt(0) === "#" ? selOrName.slice(1) : selOrName;
    return $("#" + lsId(id));
}

export function prefixHtmlIds(html, prefix) {
    if (!html || !prefix) {
        return html;
    }
    return String(html).replace(/\b(id|for|aria-controls)="([^"]*)"/g, function (match, attr, id) {
        if (!isWidgetId(id) || id.indexOf(prefix) === 0) {
            return match;
        }
        return attr + '="' + prefix + id + '"';
    });
}

export function scopeSelector(selector) {
    if (typeof selector !== "string" || selector.charAt(0) !== "#") {
        return selector;
    }
    const prefix = currentPrefix();
    if (!prefix) {
        return selector;
    }
    const hashEnd = selector.search(/[\s\[\.:,>+~]/);
    const id = hashEnd === -1 ? selector.slice(1) : selector.slice(1, hashEnd);
    const rest = hashEnd === -1 ? "" : selector.slice(hashEnd);
    if (!isWidgetId(id) || id.indexOf(prefix) === 0) {
        return selector;
    }
    return "#" + prefix + id + rest;
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

export function installDomScopeHooks() {
    if (typeof window === "undefined" || installDomScopeHooks.done) {
        return;
    }
    const $ = window.jQuery;
    if (!$) {
        return;
    }
    installDomScopeHooks.done = true;

    const rawInit = $.fn.init;
    $.fn.init = function (selector, context, root) {
        if (typeof selector === "string" && selector.charAt(0) === "#" && context == null) {
            selector = scopeSelector(selector);
        }
        return rawInit.call(this, selector, context, root);
    };
    $.fn.init.prototype = rawInit.prototype;

    ["append", "prepend", "before", "after"].forEach(function (name) {
        const raw = $.fn[name];
        $.fn[name] = function () {
            const prefix = currentPrefix();
            const args = Array.prototype.slice.call(arguments).map(function (arg) {
                return typeof arg === "string" ? prefixHtmlIds(arg, prefix) : arg;
            });
            return raw.apply(this, args);
        };
    });

    const rawHtml = $.fn.html;
    $.fn.html = function (value) {
        if (arguments.length && typeof value === "string") {
            return rawHtml.call(this, prefixHtmlIds(value, currentPrefix()));
        }
        return rawHtml.apply(this, arguments);
    };

    if (window.Document && Document.prototype.getElementById && !Document.prototype.getElementById.__lsScoped) {
        const rawGet = Document.prototype.getElementById;
        Document.prototype.getElementById = function (id) {
            if (isWidgetId(id)) {
                const prefix = currentPrefix();
                if (prefix && id.indexOf(prefix) !== 0) {
                    const scoped = rawGet.call(this, prefix + id);
                    if (scoped) {
                        return scoped;
                    }
                }
            }
            return rawGet.call(this, id);
        };
        Document.prototype.getElementById.__lsScoped = true;
    }
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
    el.addEventListener("pointerdown", function () {
        const id = el.getAttribute("data-ls-host");
        if (!id) {
            return;
        }
        if (getFocusedId() !== id) {
            focusHook(id);
        }
    }, true);
}
