package com.xc.luckysheet.server.mutation;

import com.xc.luckysheet.server.contract.WorkbookAclRole;

/** Common server-owned policy metadata for JSON workbook reducers. */
abstract class CanonicalJsonMutationDescriptor implements MutationDescriptor {
    private final String id;
    private final WorkbookAclRole requiredRole;
    private final boolean checksProtection;
    private final String protectionAction;

    CanonicalJsonMutationDescriptor(
            String id,
            WorkbookAclRole requiredRole,
            boolean checksProtection,
            String protectionAction
    ) {
        this.id = id;
        this.requiredRole = requiredRole;
        this.checksProtection = checksProtection;
        this.protectionAction = protectionAction;
    }

    @Override
    public final String id() {
        return id;
    }

    @Override
    public boolean internalOnly() {
        return false;
    }

    @Override
    public final WorkbookAclRole requiredRole() {
        return requiredRole;
    }

    @Override
    public final MutationRebasePolicy rebasePolicy() {
        return MutationRebasePolicy.EXACT_BASE;
    }

    @Override
    public final boolean checksProtection() {
        return checksProtection;
    }

    @Override
    public final String protectionAction() {
        return protectionAction;
    }
}
