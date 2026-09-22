package com.xc.luckysheet.server.mutation;

import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.contract.GeneratedWorkbookContract;

/** Common server-owned policy metadata for JSON workbook reducers. */
abstract class CanonicalJsonMutationDescriptor implements MutationDescriptor {
    private final String id;
    private final WorkbookAclRole requiredRole;
    private final boolean checksProtection;
    private final String protectionAction;

    CanonicalJsonMutationDescriptor(String id, WorkbookAclRole requiredRole) {
        GeneratedWorkbookContract.PermissionPolicy permission = GeneratedWorkbookContract.mutationPermission(id);
        if (permission == null) throw new IllegalStateException("Mutation is missing a generated permission policy: " + id);
        this.id = id;
        this.requiredRole = requiredRole;
        this.checksProtection = permission.checksProtection();
        this.protectionAction = permission.protectionAction();
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
