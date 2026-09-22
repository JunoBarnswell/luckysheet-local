package com.xc.luckysheet.server.migration;

import com.fasterxml.jackson.databind.ObjectMapper;
import org.flywaydb.core.api.migration.BaseJavaMigration;
import org.flywaydb.core.api.migration.Context;
import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;
import java.util.HexFormat;

public abstract class CanonicalSnapshotMigration extends BaseJavaMigration {
    @Override public Integer getChecksum() { return 1; }
    @Override public void migrate(Context context) throws Exception {
        var mapper = new ObjectMapper();
        var connection = context.getConnection();
        try (var select = connection.prepareStatement("select unit_id, revision, snapshot_json, checksum from snapshot_checkpoint");
             var rows = select.executeQuery();
             var update = connection.prepareStatement("update snapshot_checkpoint set snapshot_json=?, checksum=? where unit_id=? and revision=?")) {
            while (rows.next()) {
                String oldJson = rows.getString(3);
                if (!checksum(oldJson).equals(rows.getString(4))) throw new IllegalStateException("CHECKPOINT_CHECKSUM_MISMATCH: " + rows.getString(1) + "/" + rows.getLong(2));
                String json = mapper.writeValueAsString(SnapshotUpgrade.migrateStored(mapper.readTree(oldJson), rows.getString(1)));
                update.setString(1, json); update.setString(2, checksum(json)); update.setString(3, rows.getString(1)); update.setLong(4, rows.getLong(2)); update.executeUpdate();
            }
        }
        try (var select = connection.prepareStatement("select unit_id, snapshot_json from workbooks"); var rows = select.executeQuery();
             var update = connection.prepareStatement("update workbooks set snapshot_json=? where unit_id=?")) {
            while (rows.next()) {
                String json = mapper.writeValueAsString(SnapshotUpgrade.migrateStored(mapper.readTree(rows.getString(2)), rows.getString(1)));
                update.setString(1, json); update.setString(2, rows.getString(1)); update.executeUpdate();
            }
        }
    }
    private String checksum(String json) throws Exception {
        return HexFormat.of().formatHex(MessageDigest.getInstance("SHA-256").digest(json.getBytes(StandardCharsets.UTF_8)));
    }
}
