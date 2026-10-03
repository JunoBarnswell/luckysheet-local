package com.xc.luckysheet.server.service;

import jakarta.persistence.EntityManager;
import jakarta.persistence.PersistenceContext;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

/** One capacity authority for blocks, images and native artifacts. Totals are
 * derived from committed rows, so rollback, replacement and cascading purge
 * cannot leave a separate usage ledger out of sync. The existing quota row
 * serializes capacity checks and writes across workbooks and server instances. */
@Service
public class WorkbookResourceQuotaService {
    public static final long MAX_GLOBAL_BYTES = 1024L * 1024 * 1024;
    public static final long MAX_WORKBOOK_BYTES = 256L * 1024 * 1024;
    public static final long MAX_OWNER_BYTES = 512L * 1024 * 1024;
    public static final long MAX_SPACE_BYTES = 768L * 1024 * 1024;
    @PersistenceContext private EntityManager em;
    private static final String RESOURCES = "(select unit_id, byte_length from workbook_asset union all "
            + "select unit_id, byte_length from workbook_source_artifact union all "
            + "select unit_id, byte_length from workbook_data_block) r";

    @Transactional(propagation = Propagation.MANDATORY)
    public void requireCapacity(String unitId, long additionalBytes, long additionalObjects) {
        em.createNativeQuery("select id from workbook_resource_quota where id = 1 for update").getSingleResult();
        // Flush pending copies/deletes before counting; never count binary payloads.
        em.flush();
        check("Server", usage("", null), additionalBytes, additionalObjects, MAX_GLOBAL_BYTES, 40_000);
        check("Workbook", usage(" where r.unit_id = :scope", unitId), additionalBytes, additionalObjects, MAX_WORKBOOK_BYTES, 10_000);
        for (String field : java.util.List.of("owner_subject", "space_id")) {
            String scope = (String) em.createNativeQuery("select " + field + " from workbooks where unit_id = :unit")
                    .setParameter("unit", unitId).getSingleResult();
            check(field.equals("owner_subject") ? "Owner" : "Space",
                    usage(" join workbooks w on w.unit_id = r.unit_id where w." + field + (scope == null ? " is null" : " = :scope"), scope),
                    additionalBytes, additionalObjects, field.equals("owner_subject") ? MAX_OWNER_BYTES : MAX_SPACE_BYTES,
                    field.equals("owner_subject") ? 20_000 : 30_000);
        }
    }

    @Transactional(propagation = Propagation.MANDATORY)
    public void requireSpaceDestinationCapacity(String unitId, String targetSpaceId) {
        em.createNativeQuery("select id from workbook_resource_quota where id = 1 for update").getSingleResult();
        em.flush();
        String currentSpace = (String) em.createNativeQuery("select space_id from workbooks where unit_id = :unit")
                .setParameter("unit", unitId).getSingleResult();
        if (java.util.Objects.equals(currentSpace, targetSpaceId)) return;
        Object[] workbook = usage(" where r.unit_id = :scope", unitId);
        Object[] target = usage(" join workbooks w on w.unit_id = r.unit_id where w.space_id = :scope", targetSpaceId);
        check("Destination space", target, ((Number) workbook[0]).longValue(), ((Number) workbook[1]).longValue(), MAX_SPACE_BYTES, 30_000);
    }

    private Object[] usage(String suffix, String scope) {
        var query = em.createNativeQuery("select coalesce(sum(r.byte_length), 0), count(*) from " + RESOURCES + suffix);
        if (scope != null) query.setParameter("scope", scope);
        return (Object[]) query.getSingleResult();
    }
    private void check(String scope, Object[] usage, long bytes, long count, long maxBytes, long maxCount) {
        long usedBytes = ((Number) usage[0]).longValue(), usedCount = ((Number) usage[1]).longValue();
        if (bytes > 0 && usedBytes > maxBytes - bytes || count > 0 && usedCount > maxCount - count) {
            throw new ServiceException("STORAGE_QUOTA_EXCEEDED", 413,
                    scope + " storage quota exceeded; release unreferenced resources or purge unused workbooks before retrying");
        }
    }
}
