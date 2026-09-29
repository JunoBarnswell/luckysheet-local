package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.xc.luckysheet.server.contract.OperationMutation;

/** Internal capability for reducers operating on a detached, transaction-owned snapshot. */
interface OwnedSnapshotMutationDescriptor {
    /** Mutates the exclusively owned detached snapshot and returns the same root identity. */
    MutationApplication applyWithPatchOnOwnedSnapshot(JsonNode ownedSnapshot, OperationMutation mutation);
}
