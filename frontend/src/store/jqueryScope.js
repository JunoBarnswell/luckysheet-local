import { lsId, prefixHtmlIds, scopeSelector, markInstanceNodes } from "./domScope";
import { getFocusedContext, getFocusedId, getUnit, withInstance } from "./registry";

function getJQuery() {
    if (typeof window === "undefined" || !window.jQuery) {
        throw new Error("jQuery is unavailable");
    }
    return window.jQuery;
}

function normalizeSelector(selector) {
    if (typeof selector !== "string") {
        return selector;
    }
    const text = selector.trim();
    if (text.charAt(0) === "<") {
        return prefixHtmlIds(selector);
    }
    return scopeSelector(selector);
}

function normalizeMarkup(value) {
    return typeof value === "string" ? prefixHtmlIds(value) : value;
}

function ownerForRegistration() {
    const context = getFocusedContext();
    return context ? context.instanceId : null;
}

function scopeEventNames(value, instanceId) {
    if (!instanceId || typeof value !== "string") {
        return value;
    }
    return value.split(/\s+/).map(function (item) {
        if (item.indexOf("luckysheet") === -1 || item.indexOf("." + instanceId) > -1) {
            return item;
        }
        return item + "." + instanceId;
    }).join(" ");
}

function belongsToInstance(event, instanceId) {
    if (!event || !event.target) {
        return getFocusedId() === instanceId;
    }
    const context = getUnit(instanceId);
    if (!context) {
        return false;
    }
    const target = event.target.nodeType === 1 ? event.target : event.target.parentElement;
    if (!target) {
        return getFocusedId() === instanceId;
    }
    const owned = target.closest && target.closest('[data-ls-instance="' + instanceId + '"]');
    return !!owned || target.closest("#" + context.container) != null || getFocusedId() === instanceId;
}

function bindInstanceHandler(handler, instanceId) {
    if (!instanceId || typeof handler !== "function") {
        return handler;
    }
    return function () {
        const event = arguments[0];
        if (!getUnit(instanceId) || !belongsToInstance(event, instanceId)) {
            return undefined;
        }
        const receiver = this;
        const args = arguments;
        return withInstance(instanceId, function () {
            return handler.apply(receiver, args);
        });
    };
}

function registerInsertedPortals(target, instanceId) {
    const context = getUnit(instanceId);
    if (!context) {
        return;
    }
    const selector = context.domPrefix
        ? '[id^="' + context.domPrefix + '"]'
        : '[id^="luckysheet"], [id^="testdpidiv"], [id^="cellDatePickerBtn"]';
    const nodes = target.find(selector).add(target.filter(selector));
    markInstanceNodes(nodes, instanceId);
    if (target.is && target.is("body")) {
        nodes.each(function () {
            if (this.parentElement === document.body) {
                context.portals.add(this);
            }
        });
    }
}

function wrapCollection(collection) {
    let facade;
    facade = new Proxy(collection, {
        get: function (target, property) {
            const value = target[property];
            if (typeof value !== "function") {
                return value;
            }
            if (["append", "prepend", "before", "after", "html", "replaceWith"].indexOf(property) > -1) {
                return function () {
                    if (property === "html" && arguments.length === 0) {
                        return value.call(target);
                    }
                    const instanceId = ownerForRegistration();
                    const args = Array.prototype.slice.call(arguments).map(normalizeMarkup);
                    const result = value.apply(target, args);
                    registerInsertedPortals(target, instanceId);
                    return result === target ? facade : (result && result.jquery ? wrapCollection(result) : result);
                };
            }
            if (property === "on" || property === "one" || property === "off") {
                return function () {
                    const instanceId = ownerForRegistration();
                    const args = Array.prototype.slice.call(arguments);
                    if (typeof args[0] === "string") {
                        args[0] = scopeEventNames(args[0], instanceId);
                    }
                    if (property !== "off") {
                        for (let index = args.length - 1; index >= 0; index--) {
                            if (typeof args[index] === "function") {
                                args[index] = bindInstanceHandler(args[index], instanceId);
                                break;
                            }
                        }
                    }
                    const result = value.apply(target, args);
                    return result === target ? facade : (result && result.jquery ? wrapCollection(result) : result);
                };
            }
            if (["find", "closest", "filter", "is", "not"].indexOf(property) > -1) {
                return function () {
                    const args = Array.prototype.slice.call(arguments);
                    if (typeof args[0] === "string") {
                        args[0] = normalizeSelector(args[0]);
                    }
                    const result = value.apply(target, args);
                    return result === target ? facade : (result && result.jquery ? wrapCollection(result) : result);
                };
            }
            return function () {
                const result = value.apply(target, arguments);
                return result === target ? facade : (result && result.jquery ? wrapCollection(result) : result);
            };
        },
    });
    return facade;
}

function scopedJQueryImpl(selector, context) {
    const $ = getJQuery();
    const collection = context == null ? $(normalizeSelector(selector)) : $(normalizeSelector(selector), context);
    return wrapCollection(collection);
}

/**
 * Drop-in jQuery import for LuckySheet modules. Static jQuery helpers are
 * forwarded, while instance-owned IDs are resolved through InstanceDom.
 */
const scopedJQuery = new Proxy(scopedJQueryImpl, {
    apply: function (_, __, args) {
        return scopedJQueryImpl(args[0], args[1]);
    },
    get: function (_, property) {
        const value = getJQuery()[property];
        return typeof value === "function" ? value.bind(getJQuery()) : value;
    },
});

export { lsId };
export default scopedJQuery;
