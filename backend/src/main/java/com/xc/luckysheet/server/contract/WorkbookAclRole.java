package com.xc.luckysheet.server.contract;

public enum WorkbookAclRole {
    OWNER(4),
    EDITOR(3),
    COMMENTER(2),
    VIEWER(1);

    private final int rank;

    WorkbookAclRole(int rank) {
        this.rank = rank;
    }

    public boolean includes(WorkbookAclRole required) {
        return rank >= required.rank;
    }
}
