package com.xc.luckysheet.server.service;

import com.xc.luckysheet.server.contract.WorkbookRole;

import java.util.List;

public record RangeAccessContext(String subject, WorkbookRole workbookRole, List<String> groups, long accessRevision) {
    public RangeAccessContext {
        if (subject == null || subject.isBlank() || workbookRole == null || accessRevision < 0) {
            throw new IllegalArgumentException("Range access context is invalid");
        }
        groups = groups == null ? List.of() : groups.stream().filter(value -> value != null && !value.isBlank()).distinct().toList();
    }
}
