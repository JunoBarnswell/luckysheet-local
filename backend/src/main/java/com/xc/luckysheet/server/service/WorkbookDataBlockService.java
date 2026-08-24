package com.xc.luckysheet.server.service;

import com.xc.luckysheet.server.contract.DataBlockMetadata;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.store.DataBlockRow;
import com.xc.luckysheet.server.store.WorkbookDataBlockStore;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.io.InputStream;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.time.Instant;
import java.util.HexFormat;

@Service
public class WorkbookDataBlockService {
    /** A single block must remain bounded; manifests split large sources. */
    public static final int MAX_BLOCK_BYTES = 32 * 1024 * 1024;
    public static final long MAX_WORKBOOK_BLOCK_BYTES = 256L * 1024 * 1024;
    public static final long MAX_WORKBOOK_BLOCK_COUNT = 10_000;

    private final WorkbookDataBlockStore store;
    private final AccessControlService access;
    private final WorkbookLifecycleService lifecycle;

    public WorkbookDataBlockService(WorkbookDataBlockStore store, AccessControlService access, WorkbookLifecycleService lifecycle) {
        this.store = store;
        this.access = access;
        this.lifecycle = lifecycle;
    }

    public DataBlockMetadata put(String unitId, String sourceId, String blockId, String checksum, long contentLength,
                                 InputStreamSource contentSource, String actor) {
        access.require(unitId, actor, WorkbookAclRole.EDITOR);
        lifecycle.requireActive(unitId);
        validateIdentity(sourceId, "sourceId");
        validateIdentity(blockId, "blockId");
        if (contentLength == 0 || contentLength > MAX_BLOCK_BYTES) {
            throw ServiceException.validation("Data block size must be between 1 and " + MAX_BLOCK_BYTES + " bytes");
        }
        byte[] content = readBounded(contentSource);
        String actualChecksum = sha256(content);
        if (!actualChecksum.equals(checksum)) throw ServiceException.validation("Data block checksum mismatch");
        Instant now = Instant.now();
        DataBlockRow row = new DataBlockRow(unitId, sourceId, blockId, actualChecksum, content.length, content.clone(), now, now);
        return metadata(store.upsertWithinQuota(row, MAX_WORKBOOK_BLOCK_BYTES, MAX_WORKBOOK_BLOCK_COUNT));
    }

    public DataBlockRow get(String unitId, String sourceId, String blockId, String actor) {
        access.require(unitId, actor, WorkbookAclRole.VIEWER);
        lifecycle.requireActive(unitId);
        validateIdentity(sourceId, "sourceId");
        validateIdentity(blockId, "blockId");
        return store.find(unitId, sourceId, blockId)
                .orElseThrow(() -> ServiceException.notFound("Data block not found"));
    }

    public void delete(String unitId, String sourceId, String blockId, String actor) {
        access.require(unitId, actor, WorkbookAclRole.EDITOR);
        lifecycle.requireActive(unitId);
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

    private static byte[] readBounded(InputStreamSource source) {
        try (InputStream input = source.open()) {
            byte[] content = input.readNBytes(MAX_BLOCK_BYTES + 1);
            if (content.length == 0 || content.length > MAX_BLOCK_BYTES) {
                throw ServiceException.validation("Data block size must be between 1 and " + MAX_BLOCK_BYTES + " bytes");
            }
            return content;
        } catch (IOException error) {
            throw new ServiceException("VALIDATION_ERROR", 400, "Data block body could not be read", error);
        }
    }

    @FunctionalInterface
    public interface InputStreamSource {
        InputStream open() throws IOException;
    }
}
