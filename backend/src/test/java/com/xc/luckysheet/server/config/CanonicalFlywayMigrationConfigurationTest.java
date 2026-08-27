package com.xc.luckysheet.server.config;

import org.flywaydb.core.Flyway;
import org.flywaydb.core.api.MigrationInfo;
import org.flywaydb.core.api.MigrationInfoService;
import org.flywaydb.core.api.MigrationState;
import org.flywaydb.core.api.MigrationVersion;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.mockito.Mockito.inOrder;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.when;

class CanonicalFlywayMigrationConfigurationTest {
    @Test
    void retiresOnlyTheRecordedRowsPermutedMigrationBeforeRunningTheCanonicalOwner() {
        Flyway flyway = mock(Flyway.class);
        MigrationInfoService before = mock(MigrationInfoService.class);
        MigrationInfoService after = mock(MigrationInfoService.class);
        MigrationInfo retired = migration("5", "canonical rows permuted ranges", MigrationState.MISSING_SUCCESS);
        MigrationInfo deleted = migration("5", "canonical rows permuted ranges", MigrationState.DELETED);
        when(flyway.info()).thenReturn(before, after);
        when(before.all()).thenReturn(new MigrationInfo[]{retired});
        when(after.all()).thenReturn(new MigrationInfo[]{deleted});

        CanonicalFlywayMigrationConfiguration.migrate(flyway);

        var ordered = inOrder(flyway);
        ordered.verify(flyway).repair();
        ordered.verify(flyway).migrate();
    }

    @Test
    void rejectsAnyOtherMissingMigrationWithoutRepairingHistory() {
        Flyway flyway = mock(Flyway.class);
        MigrationInfoService info = mock(MigrationInfoService.class);
        MigrationInfo unrelated = migration("4", "another migration", MigrationState.MISSING_SUCCESS);
        when(flyway.info()).thenReturn(info);
        when(info.all()).thenReturn(new MigrationInfo[]{unrelated});

        assertThrows(IllegalStateException.class, () -> CanonicalFlywayMigrationConfiguration.migrate(flyway));
    }

    private static MigrationInfo migration(String version, String description, MigrationState state) {
        MigrationInfo migration = mock(MigrationInfo.class);
        when(migration.getVersion()).thenReturn(MigrationVersion.fromVersion(version));
        when(migration.getDescription()).thenReturn(description);
        when(migration.getState()).thenReturn(state);
        return migration;
    }
}
