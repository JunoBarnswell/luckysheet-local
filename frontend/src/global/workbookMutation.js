import Store from "../store";

function mutationState() {
    Store.runtime.mutations = Store.runtime.mutations || { sequence: 0, last: null };
    return Store.runtime.mutations;
}

export function createOperationId() {
    const state = mutationState();
    state.sequence += 1;
    return [Store.instanceId || "legacy", Date.now().toString(36), state.sequence.toString(36)].join(":");
}

export function beginWorkbookMutation(type, sheetIndex) {
    return {
        operationId: createOperationId(),
        type: type,
        sheetIndex: sheetIndex == null ? Store.currentSheetIndex : sheetIndex,
        changedCells: [],
        affectedSheets: {},
        dirtyRegions: [],
        historyPatch: [],
    };
}

export function recordChangedCell(mutation, row, column, before, after, sheetIndex) {
    if (!mutation || row == null || column == null) return mutation;
    const index = sheetIndex == null ? mutation.sheetIndex : sheetIndex;
    mutation.changedCells.push({ r: row, c: column, index: index, before: before, after: after });
    mutation.affectedSheets[index] = true;
    mutation.dirtyRegions.push({ row: [row, row], column: [column, column] });
    mutation.historyPatch.push({ r: row, c: column, index: index, before: before, after: after });
    return mutation;
}

export function commitWorkbookMutation(mutation) {
    if (!mutation) return null;
    mutation.affectedSheets = Object.keys(mutation.affectedSheets).map(function (index) { return String(index); });
    mutationState().last = mutation;
    return mutation;
}

export function getLastWorkbookMutation() {
    return mutationState().last;
}
