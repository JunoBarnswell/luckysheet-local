import { requireFocusedContext } from "./registry";

/**
 * Store 默认导出是 Proxy → focused 实例上下文。
 * 禁止 `const { flowdata } = Store` 后长期持有；必须每次从 Store 读。
 */
function target() {
    return requireFocusedContext();
}

const Store = new Proxy({}, {
    get(_, prop) {
        if (prop === "__ctx") {
            return target();
        }
        const ctx = target();
        const value = ctx[prop];
        return typeof value === "function" ? value.bind(ctx) : value;
    },
    set(_, prop, value) {
        target()[prop] = value;
        return true;
    },
    has(_, prop) {
        return prop in target();
    },
    deleteProperty(_, prop) {
        return delete target()[prop];
    },
    ownKeys() {
        return Reflect.ownKeys(target());
    },
    getOwnPropertyDescriptor(_, prop) {
        const ctx = target();
        if (!(prop in ctx)) {
            return undefined;
        }
        return {
            configurable: true,
            enumerable: true,
            writable: true,
            value: ctx[prop],
        };
    },
});

export default Store;
