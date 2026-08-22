export const RESOURCE_WAIT_MS = 10000;

/** Drains dynamically registered print resources until the deadline. */
export class PrintResourceCollector {
    constructor() {
        this.pending = new Set();
        this.diagnostics = [];
    }

    add(resource, label) {
        const tracked = Promise.resolve(resource)
            .catch((error) => {
                this.diagnostics.push({ label: label || "resource", error: String(error && error.message || error) });
            })
            .then(() => {
                this.pending.delete(tracked);
            });
        this.pending.add(tracked);
        return tracked;
    }

    wait(timeout) {
        const deadline = Date.now() + (timeout == null ? RESOURCE_WAIT_MS : timeout);
        const self = this;
        function drain() {
            if (!self.pending.size) {
                return Promise.resolve({ timedOut: false, diagnostics: self.diagnostics.slice() });
            }
            const remaining = deadline - Date.now();
            if (remaining <= 0) {
                return Promise.resolve({ timedOut: true, diagnostics: self.diagnostics.slice() });
            }
            return Promise.race([
                Promise.all(Array.from(self.pending)),
                new Promise((resolve) => setTimeout(resolve, remaining)),
            ]).then(drain);
        }
        return drain();
    }
}
