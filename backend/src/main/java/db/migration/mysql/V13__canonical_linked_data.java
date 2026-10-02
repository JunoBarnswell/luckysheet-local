package db.migration.mysql;
public class V13__canonical_linked_data extends com.xc.luckysheet.server.migration.CanonicalSnapshotMigration {
    @Override public void migrate(org.flywaydb.core.api.migration.Context context) throws Exception {
        super.migrate(context);
        new R__structural_patch_v2().migrate(context);
    }
}
