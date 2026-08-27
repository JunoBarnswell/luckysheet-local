package com.xc.luckysheet.server.config;

import org.flywaydb.core.Flyway;
import org.flywaydb.core.api.MigrationInfo;
import org.flywaydb.core.api.MigrationState;
import org.springframework.boot.autoconfigure.flyway.FlywayMigrationStrategy;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

/**
 * Owns the one-time clean break from the retired numeric Java migration to the
 * canonical repeatable migration. No runtime business path depends on legacy
 * envelopes after this startup boundary completes.
 */
@Configuration
public class CanonicalFlywayMigrationConfiguration {
    private static final String RETIRED_VERSION = "5";
    private static final String RETIRED_DESCRIPTION = "canonical rows permuted ranges";

    @Bean
    FlywayMigrationStrategy canonicalFlywayMigrationStrategy() {
        return CanonicalFlywayMigrationConfiguration::migrate;
    }

    static void migrate(Flyway flyway) {
        boolean retiredMigrationRecorded = false;
        for (MigrationInfo migration : flyway.info().all()) {
            MigrationState state = migration.getState();
            if (state == MigrationState.MISSING_SUCCESS && isRetiredMigration(migration)) {
                retiredMigrationRecorded = true;
                continue;
            }
            if (state == MigrationState.MISSING_SUCCESS || state == MigrationState.MISSING_FAILED
                    || state == MigrationState.FAILED || state == MigrationState.FUTURE_FAILED) {
                throw new IllegalStateException("Flyway migration history is not canonical: " + describe(migration));
            }
            if (state.isApplied() && state.isResolved()
                    && (!migration.isChecksumMatching() || !migration.isDescriptionMatching())) {
                throw new IllegalStateException("Flyway migration history does not match the resolved migration: " + describe(migration));
            }
        }
        if (retiredMigrationRecorded) {
            flyway.repair();
            for (MigrationInfo migration : flyway.info().all()) {
                if (migration.getState() == MigrationState.MISSING_SUCCESS || migration.getState() == MigrationState.MISSING_FAILED) {
                    throw new IllegalStateException("Flyway clean-break repair left a missing migration: " + describe(migration));
                }
            }
        }
        flyway.migrate();
    }

    private static boolean isRetiredMigration(MigrationInfo migration) {
        return migration.getVersion() != null
                && RETIRED_VERSION.equals(migration.getVersion().getVersion())
                && RETIRED_DESCRIPTION.equals(migration.getDescription());
    }

    private static String describe(MigrationInfo migration) {
        return "version=" + migration.getVersion() + ", description=" + migration.getDescription() + ", state=" + migration.getState();
    }
}
