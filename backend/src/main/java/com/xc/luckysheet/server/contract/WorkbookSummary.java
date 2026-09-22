package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonInclude;

import java.time.Instant;
import java.util.List;

/**
 * One list item contains all actor-dependent metadata needed by the Hub.
 * The actor role is calculated by the server from direct ACL and space
 * membership; clients never submit or infer it.
 */
@JsonInclude(JsonInclude.Include.NON_NULL)
public record WorkbookSummary(
        String unitId,
        String name,
        long revision,
        Instant updatedAt,
        WorkbookAclRole role,
        String ownerSubject,
        String spaceId,
        String folderId,
        List<String> locationPath,
        String spaceName,
        String sourceFileName,
        WorkbookStorageLocation storageLocation,
        WorkbookSyncStatus syncStatus,
        WorkbookLifecycle lifecycle,
        WorkbookSource source,
        boolean favorite,
        Instant lastOpenedAt,
        Instant deletedAt
) {
    /** Existing operation tests and non-catalog callers retain a valid item. */
    public WorkbookSummary(String unitId, String name, long revision, Instant updatedAt) {
        this(unitId, name, revision, updatedAt, null, null, null, null, null, null, null,
                WorkbookStorageLocation.REMOTE, WorkbookSyncStatus.SYNCED,
                WorkbookLifecycle.ACTIVE, WorkbookSource.NATIVE, false, null, null);
    }
}
