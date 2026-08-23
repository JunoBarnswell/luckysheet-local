package com.xc.luckysheet.server.service;

import com.xc.luckysheet.server.contract.DataBlockMetadata;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.store.DataBlockRow;
import com.xc.luckysheet.server.store.WorkbookDataBlockStore;
import org.springframework.stereotype.Service;

import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.time.Instant;
import java.util.HexFormat;

@Service
public class WorkbookDataBlockService {
    /** A single block must remain bounded; manifests split large sources. */
    public static final int MAX_BLOCK_BYTES = 32 * 1024 * 1024;

    private final WorkbookDataBlockStore store;
    private final AccessControlService access;

    public WorkbookDataBlockService(WorkbookDataBlockStore store, AccessControlService access) {
        this.store = store;
        this.access = access;
    }

    public DataBlockMetadata put(String unitId, String sourceId, String blockId, String checksum, byte[] content, String actor) {
        access.require(unitId, actor, WorkbookAclRole.EDITOR);
        validateIdentity(sourceId, "sourceId");
        validateIdentity(blockId, "blockId");
        if (content == null || content.length == 0 || content.length > MAX_BLOCK_BYTES) {
            throw ServiceException.validation("Data block size must be between 1 and " + MAX_BLOCK_BYTES + " bytes");
        }
        String actualChecksum = sha256(content);
        if (!actualChecksum.equals(checksum)) throw ServiceException.validation("Data block checksum mismatch");
        Instant now = Instant.now();
        DataBlockRow existing = store.find(unitId, sourceId, blockId).orElse(null);
        if (existing != null && existing.checksum().equals(actualChecksum) && existing.byteLength() == content.length) {
            return metadata(existing);
        }
        DataBlockRow row = new DataBlockRow(unitId, sourceId, blockId, actualChecksum, content.length, content.clone(), now, now);
        store.upsert(row);
        return metadata(row);
    }

    public DataBlockRow get(String unitId, String sourceId, String blockId, String actor) {
        access.require(unitId, actor, WorkbookAclRole.VIEWER);
        validateIdentity(sourceId, "sourceId");
        validateIdentity(blockId, "blockId");
        return store.find(unitId, sourceId, blockId)
                .orElseThrow(() -> ServiceException.notFound("Data block not found"));
    }

    public void delete(String unitId, String sourceId, String blockId, String actor) {
        access.require(unitId, actor, WorkbookAclRole.EDITOR);
        validateIdentity(sourceId, "sourceId");
        validateIdentity(blockId, "blockId");
        store.delete(unitId, sourceId, blockId);
    }

    private static DataBlockMetadata metadata(DataBlockRow row) {
        return new DataBlockMetadata(row.unitId(), row.sourceId(), row.blockId(), row.checksum(), row.byteLength(), row.updatedAt());
    }

    private static void validateIdentity(String value, String field) {
        if (value == null || !value.matches("[A-Za-z0-9._:-]{1,200}")) {
            throw ServiceException.validation(field + " is invalid");
        }
    }

    private static String sha256(byte[] content) {
        try {
            return HexFormat.of().formatHex(MessageDigest.getInstance("SHA-256").digest(content));
        } catch (NoSuchAlgorithmException error) {
            throw new IllegalStateException("SHA-256 is unavailable", error);
        }
    }
}
