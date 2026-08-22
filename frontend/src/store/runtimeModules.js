import { getFocusedContext } from "./registry";

function cloneRuntimeValue(value) {
    if (Array.isArray(value)) {
        return value.map(cloneRuntimeValue);
    }
    if (value && Object.prototype.toString.call(value) === "[object Object]") {
        const copy = {};
        Object.keys(value).forEach(function (key) {
            copy[key] = cloneRuntimeValue(value[key]);
        });
        return copy;
    }
    return value;
}

function getRuntime(ctx, name, template) {
    if (!ctx || ctx.disposed) {
        throw new Error("LuckySheet instance runtime is not available");
    }
    ctx.modules = ctx.modules || {};
    if (!ctx.modules[name]) {
        ctx.modules[name] = cloneRuntimeValue(template);
    }
    return ctx.modules[name];
}

/**
 * Turns a legacy exported object into a context-owned runtime object. Every
 * property read/write resolves against the focused workbook context, while the
 * object identity imported by old modules remains stable.
 */
export function createContextualModule(name, template) {
    let facade;
    facade = new Proxy(template, {
        get: function (_, property) {
            const value = getRuntime(getFocusedContext(), name, template)[property];
            return typeof value === "function" ? value.bind(facade) : value;
        },
        set: function (_, property, value) {
            getRuntime(getFocusedContext(), name, template)[property] = value;
            return true;
        },
        has: function (_, property) {
            return property in getRuntime(getFocusedContext(), name, template);
        },
        ownKeys: function () {
            return Reflect.ownKeys(getRuntime(getFocusedContext(), name, template));
        },
        getOwnPropertyDescriptor: function (_, property) {
            const runtime = getRuntime(getFocusedContext(), name, template);
            if (!(property in runtime)) {
                return undefined;
            }
            return {
                configurable: true,
                enumerable: true,
                writable: true,
                value: runtime[property],
            };
        },
    });
    return facade;
}

export function getRuntimeModule(context, name) {
    return context && context.modules ? context.modules[name] || null : null;
}
