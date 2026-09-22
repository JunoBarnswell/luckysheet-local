package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.contract.WorkbookAclRole;

import java.util.List;

public interface MutationDescriptor {
    String id();

    boolean internalOnly();

    /**
     * The lowest persistent workbook role that may submit this mutation.
     * Roles and affected ranges are resolved by the server from the current
     * snapshot; clients never supply either as authority.
     */
    WorkbookAclRole requiredRole();

    /** The server's concurrency rule for this concrete mutation. */
    MutationRebasePolicy rebasePolicy();

    /**
     * Whether locked workbook, sheet, or range rules apply to this mutation.
     * Administrative mutations (for example changing protection itself) are
     * intentionally checked by role but are not blocked by the rule they
     * manage.
     */
    boolean checksProtection();

    /** Canonical protection action checked against a locked rule allow-list. */
    String protectionAction();

    List<RangeRef> affectedRanges(JsonNode snapshot, OperationMutation mutation);

    JsonNode apply(JsonNode snapshot, OperationMutation mutation);
}
