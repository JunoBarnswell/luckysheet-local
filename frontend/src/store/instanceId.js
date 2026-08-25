export class InvalidInstanceIdError extends Error {
    constructor(instanceId) {
        super("Invalid LuckySheet instanceId: expected a non-empty string containing only letters, numbers, underscores, or hyphens");
        this.name = "InvalidInstanceIdError";
        this.code = "INVALID_INSTANCE_ID";
        this.instanceId = instanceId;
        this.recovery = "Provide a safe instanceId matching /^[A-Za-z0-9_-]+$/";
    }
}

export function validateInstanceId(instanceId) {
    if (typeof instanceId !== "string" || !/^[A-Za-z0-9_-]+$/.test(instanceId)) {
        throw new InvalidInstanceIdError(instanceId);
    }
    return instanceId;
}
