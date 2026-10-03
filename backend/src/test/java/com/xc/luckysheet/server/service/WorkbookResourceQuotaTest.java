package com.xc.luckysheet.server.service;

import com.xc.luckysheet.server.persistence.*;
import com.xc.luckysheet.server.contract.*;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.TestPropertySource;
import org.springframework.transaction.support.TransactionTemplate;
import org.springframework.transaction.PlatformTransactionManager;
import java.time.Instant;
import java.util.UUID;
import static org.junit.jupiter.api.Assertions.*;

@SpringBootTest
@TestPropertySource(properties = {"DATABASE_URL=jdbc:h2:mem:resource_quota;DB_CLOSE_DELAY=-1", "DATABASE_USERNAME=sa", "DATABASE_PASSWORD=",
        "JPA_DDL_AUTO=validate", "FLYWAY_BASELINE_ON_MIGRATE=false", "luckysheet.auth.mode=oidc", "AUTH_ISSUER=https://issuer.test",
        "AUTH_AUDIENCE=test", "AUTH_JWKS_URL=https://issuer.test/jwks", "COORDINATION_MULTI_INSTANCE=false", "COORDINATION_REDIS_ENABLED=false"})
class WorkbookResourceQuotaTest {
    @Autowired WorkbookResourceQuotaService quota;
    @Autowired WorkbookEntityRepository workbooks;
    @Autowired AssetEntityRepository assets;
    @Autowired WorkspaceSpaceEntityRepository spaces;
    @Autowired WorkbookSourceArtifactEntityRepository artifacts;
    @Autowired PlatformTransactionManager transactions;
    private TransactionTemplate tx() { return new TransactionTemplate(transactions); }
    private String book() {
        String id = UUID.randomUUID().toString(); var now = Instant.now();
        workbooks.save(new WorkbookEntity(id, "Quota", "{}", 0, 0, now, now)); return id;
    }
    private void asset(String id, String key, int bytes) {
        var now = Instant.now(); assets.saveAndFlush(new AssetEntity(id, key, key, "image/png", bytes, null, null, new byte[]{1}, now, now));
    }
    @Test void serializedAdmissionAllowsTheLastByteAndRejectsConcurrentOversubscription() throws Exception {
        String id = book(); tx().executeWithoutResult(s -> asset(id, "seed", (int)WorkbookResourceQuotaService.MAX_WORKBOOK_BYTES - 1));
        try (var pool = java.util.concurrent.Executors.newFixedThreadPool(2)) {
            var gate = new java.util.concurrent.CountDownLatch(1);
            var futures = new java.util.ArrayList<java.util.concurrent.Future<Boolean>>();
            for (int i = 0; i < 2; i++) { final String key = "new-" + i;
                futures.add(pool.submit(() -> { gate.await(); try {
                    tx().executeWithoutResult(s -> { quota.requireCapacity(id, 1, 1); asset(id, key, 1); }); return true;
                } catch (ServiceException exceeded) { assertEquals("STORAGE_QUOTA_EXCEEDED", exceeded.code()); return false; } }));
            }
            gate.countDown(); int success = 0; for (var future : futures) if (future.get(10, java.util.concurrent.TimeUnit.SECONDS)) success++;
            assertEquals(1, success); assertEquals(2, assets.findAllByIdUnitId(id).size());
        } finally { tx().executeWithoutResult(s -> workbooks.deleteById(id)); }
    }
    @Test void nativeAndImageBytesShareCapacityAndRollbackDeletionAndReplacementDoNotDrift() {
        String id = book(); var now = Instant.now();
        tx().executeWithoutResult(s -> {
            asset(id, "seed", (int)WorkbookResourceQuotaService.MAX_WORKBOOK_BYTES - 10);
            artifacts.saveAndFlush(new WorkbookSourceArtifactEntity(id, "file.xlsx", "application/octet-stream", "hash", 10, new byte[]{1}, "{}", now, now));
        });
        try {
            assertEquals("STORAGE_QUOTA_EXCEEDED", assertThrows(ServiceException.class,
                    () -> tx().executeWithoutResult(s -> quota.requireCapacity(id, 1, 1))).code());
            assertThrows(IllegalStateException.class, () -> tx().executeWithoutResult(s -> { artifacts.deleteById(id); quota.requireCapacity(id, 10, 1); throw new IllegalStateException("rollback"); }));
            assertTrue(artifacts.existsById(id));
            tx().executeWithoutResult(s -> quota.requireCapacity(id, 0, 0)); // same-size replacement at capacity
            tx().executeWithoutResult(s -> { artifacts.deleteById(id); quota.requireCapacity(id, 10, 1); asset(id, "replacement", 10); });
            assertFalse(artifacts.existsById(id)); assertEquals(2, assets.findAllByIdUnitId(id).size());
        } finally { tx().executeWithoutResult(s -> workbooks.deleteById(id)); }
    }

    @Test void movingExistingResourcesCannotOversubscribeTheDestinationSpace() {
        String id = UUID.randomUUID().toString(), otherIdForQuotaTest = UUID.randomUUID().toString(); var now = Instant.now();
        String sourceSpace = UUID.randomUUID().toString(), destinationSpace = UUID.randomUUID().toString();
        spaces.save(new WorkspaceSpaceEntity(sourceSpace, "Source", WorkspaceSpaceType.TEAM, "owner", now, now));
        spaces.save(new WorkspaceSpaceEntity(destinationSpace, "Destination", WorkspaceSpaceType.TEAM, "owner", now, now));
        workbooks.save(new WorkbookEntity(id, "Quota", "{}", 0, 0, now, now, "owner", sourceSpace, null,
                WorkbookStorageLocation.REMOTE, WorkbookSource.NATIVE, WorkbookLifecycle.ACTIVE, null));
        tx().executeWithoutResult(s -> asset(id, "moved-resource", 1));
        try {
            // Existing target rows can reach the space cap through previously admitted writes.
            workbooks.save(new WorkbookEntity(otherIdForQuotaTest, "Existing", "{}", 0, 0, now, now, "other", destinationSpace, null,
                    WorkbookStorageLocation.REMOTE, WorkbookSource.NATIVE, WorkbookLifecycle.ACTIVE, null));
            tx().executeWithoutResult(s -> asset(otherIdForQuotaTest, "existing", (int)WorkbookResourceQuotaService.MAX_SPACE_BYTES));
            assertEquals("STORAGE_QUOTA_EXCEEDED", assertThrows(ServiceException.class,
                    () -> tx().executeWithoutResult(s -> quota.requireSpaceDestinationCapacity(id, destinationSpace))).code());
            tx().executeWithoutResult(s -> quota.requireSpaceDestinationCapacity(id, sourceSpace));
        } finally { tx().executeWithoutResult(s -> { workbooks.deleteById(id); workbooks.deleteById(otherIdForQuotaTest); spaces.deleteById(sourceSpace); spaces.deleteById(destinationSpace); }); }
    }
}
