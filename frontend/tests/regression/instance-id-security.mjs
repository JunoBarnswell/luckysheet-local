import { InvalidInstanceIdError, validateInstanceId } from "../../src/store/instanceId.js";

function assert(condition, name) {
    if (!condition) {
        throw new Error("FAIL " + name);
    }
    console.log("PASS " + name);
}

assert(validateInstanceId("workbook_1-a") === "workbook_1-a", "safe caller-provided instanceId is accepted");
assert(validateInstanceId("lsu_m123_ab12cd").startsWith("lsu_"), "generated instanceId format is accepted");

const maliciousId = 'x\"><img src=x onerror=alert(1)>';
let rejection = null;
try {
    validateInstanceId(maliciousId);
} catch (error) {
    rejection = error;
}

assert(rejection instanceof InvalidInstanceIdError, "markup-bearing instanceId is rejected");
assert(rejection && rejection.code === "INVALID_INSTANCE_ID", "rejection exposes a stable error code");
assert(rejection && rejection.instanceId === maliciousId, "rejection identifies the affected value");
assert(rejection && rejection.recovery.indexOf("^[A-Za-z0-9_-]+$") > -1, "rejection provides recovery guidance");

console.log("instance-id-security.mjs all passed");
