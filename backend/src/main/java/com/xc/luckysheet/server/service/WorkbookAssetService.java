package com.xc.luckysheet.server.service;

import com.xc.luckysheet.server.contract.AssetMetadata;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.persistence.AssetEntity;
import com.xc.luckysheet.server.persistence.AssetEntityRepository;
import com.xc.luckysheet.server.persistence.WorkbookEntityRepository;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.io.IOException;
import java.io.InputStream;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.time.Instant;
import java.util.HexFormat;
import java.util.Set;

@Service
public class WorkbookAssetService {
    public static final int MAX_ASSET_BYTES = 32 * 1024 * 1024;

    private final AssetEntityRepository assets;
    private final WorkbookEntityRepository workbooks;
    private final AccessControlService access;
    private final WorkbookLifecycleService lifecycle;

    public WorkbookAssetService(AssetEntityRepository assets, WorkbookEntityRepository workbooks,
                                AccessControlService access, WorkbookLifecycleService lifecycle) {
        this.assets = assets;
        this.workbooks = workbooks;
        this.access = access;
        this.lifecycle = lifecycle;
    }

    @Transactional
    public AssetMetadata put(String unitId, String assetId, String checksum, String mimeType,
                             Integer width, Integer height, long contentLength, InputStreamSource source, String actor) {
        authorize(unitId, actor, WorkbookAclRole.EDITOR);
        validateIdentity(assetId, "assetId");
        validateImageMime(mimeType);
        validateDimension(width, "width");
        validateDimension(height, "height");
        byte[] content = readBounded(source, contentLength);
        String actualChecksum = sha256(content);
        if (!actualChecksum.equals(checksum) || !assetId.equals("asset-" + actualChecksum)) {
            throw ServiceException.validation("Asset checksum or identity mismatch");
        }
        AssetEntity existing = assets.findById(new AssetEntity.Id(unitId, assetId)).orElse(null);
        if (existing != null) {
            if (!existing.getContentHash().equals(actualChecksum) || existing.getByteLength() != content.length || !existing.getMimeType().equals(mimeType)) {
                throw ServiceException.conflict("Asset identity already belongs to different content");
            }
            return metadata(existing);
        }
        Instant now = Instant.now();
        AssetEntity entity = new AssetEntity(unitId, assetId, actualChecksum, mimeType, content.length, width, height, content, now, now);
        assets.save(entity);
        return metadata(entity);
    }

    @Transactional(readOnly = true)
    public AssetEntity get(String unitId, String assetId, String actor) {
        authorize(unitId, actor, WorkbookAclRole.VIEWER);
        validateIdentity(assetId, "assetId");
        return assets.findById(new AssetEntity.Id(unitId, assetId)).orElseThrow(() -> ServiceException.notFound("Asset not found"));
    }

    @Transactional
    public void release(String unitId, String assetId, String actor) {
        authorize(unitId, actor, WorkbookAclRole.EDITOR);
        validateIdentity(assetId, "assetId");
        assets.deleteById(new AssetEntity.Id(unitId, assetId));
    }

    @Transactional
    public void reconcile(String unitId, Set<String> referencedAssetIds, String actor) {
        authorize(unitId, actor, WorkbookAclRole.EDITOR);
        if (referencedAssetIds == null) throw ServiceException.validation("Referenced asset set is required");
        for (AssetEntity entity : assets.findAllByIdUnitId(unitId)) {
            if (!referencedAssetIds.contains(entity.getId().getAssetId())) assets.delete(entity);
        }
    }

    public static AssetMetadata metadata(AssetEntity entity) {
        return new AssetMetadata("AssetRef", entity.getId().getUnitId(), entity.getId().getAssetId(), entity.getContentHash(),
                entity.getMimeType(), entity.getByteLength(), entity.getWidth(), entity.getHeight(), entity.getUpdatedAt());
    }

    private void authorize(String unitId, String actor, WorkbookAclRole role) {
        workbooks.findById(unitId).orElseThrow(() -> ServiceException.notFound("Workbook not found"));
        lifecycle.requireActive(unitId);
        access.require(unitId, actor, role);
    }

    private static void validateIdentity(String value, String field) {
        if (value == null || !value.matches("[A-Za-z0-9._:-]{1,200}")) throw ServiceException.validation(field + " is invalid");
    }

    private static void validateImageMime(String value) {
        if (value == null || !value.matches("image/[A-Za-z0-9.+-]{1,80}")) throw ServiceException.validation("Only image assets are supported");
    }

    private static void validateDimension(Integer value, String field) {
        if (value != null && (value <= 0 || value > 100_000)) throw ServiceException.validation(field + " is invalid");
    }

    private static byte[] readBounded(InputStreamSource source, long contentLength) {
        if (contentLength > MAX_ASSET_BYTES) throw ServiceException.validation("Asset size is invalid");
        try (InputStream input = source.open()) {
            byte[] content = input.readNBytes(MAX_ASSET_BYTES + 1);
            if (content.length == 0 || content.length > MAX_ASSET_BYTES) throw ServiceException.validation("Asset size is invalid");
            if (contentLength >= 0 && contentLength != content.length) throw ServiceException.validation("Asset content length mismatch");
            return content;
        } catch (IOException error) {
            throw new ServiceException("VALIDATION_ERROR", 400, "Asset body could not be read", error);
        }
    }

    private static String sha256(byte[] content) {
        try {
            return HexFormat.of().formatHex(MessageDigest.getInstance("SHA-256").digest(content));
        } catch (NoSuchAlgorithmException error) {
            throw new IllegalStateException("SHA-256 is unavailable", error);
        }
    }

    @FunctionalInterface
    public interface InputStreamSource { InputStream open() throws IOException; }
}
