package com.xc.luckysheet.server.migration;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import org.flywaydb.core.api.migration.Context;

/** Explicit persisted-data upgrade. Runtime reducers retain strict canonical contracts. */
public abstract class CanonicalLegacyContractsMigration extends CanonicalSnapshotMigration {
    @Override protected ObjectNode upgrade(JsonNode value, String unitId) {
        if (!(value instanceof ObjectNode object)) throw new IllegalStateException("Stored snapshot must be an object");
        ObjectNode copy = object.deepCopy();
        SnapshotUpgrade.migrateLegacyModelContracts(copy);
        return SnapshotUpgrade.migrateStored(copy, unitId);
    }
    @Override public void migrate(Context context) throws Exception {
        rewrite(context, "select operation_id, envelope_json from operation_log", "update operation_log set envelope_json=? where operation_id=?");
        rewrite(context, "select event_id, payload_json from coordination_outbox where published_at is null", "update coordination_outbox set payload_json=? where event_id=?");
        super.migrate(context);
    }
    private void rewrite(Context context, String selectSql, String updateSql) throws Exception {
        ObjectMapper mapper = new ObjectMapper();
        try (var select = context.getConnection().prepareStatement(selectSql); var rows = select.executeQuery(); var update = context.getConnection().prepareStatement(updateSql)) {
            while (rows.next()) {
                JsonNode value = mapper.readTree(rows.getString(2));
                if (!upgradeOperations(value)) continue;
                update.setString(1, mapper.writeValueAsString(value)); update.setString(2, rows.getString(1)); update.executeUpdate();
            }
        }
    }
    public static boolean upgradeOperations(JsonNode root) {
        boolean changed = false;
        java.util.ArrayDeque<JsonNode> pending = new java.util.ArrayDeque<>(); pending.add(root); int nodes = 0;
        while (!pending.isEmpty()) {
            JsonNode value = pending.pop(); if (++nodes > 2000000) throw new IllegalStateException("Stored operation exceeds migration budget");
            if (value instanceof ObjectNode mutation && mutation.get("params") instanceof ObjectNode params) {
                String id = mutation.path("id").asText();
                if (("sheetTable.add".equals(id) || "sheetTable.update".equals(id)) && params.get("table") instanceof ObjectNode table) {
                    if (!table.has("showFirstColumn")) { table.put("showFirstColumn", false); changed = true; }
                    if (!table.has("showLastColumn")) { table.put("showLastColumn", false); changed = true; }
                    if (!table.has("autoExpand")) { table.put("autoExpand", true); changed = true; }
                }
                if ("range.paste".equals(id) && !params.has("transfer")) {
                    if (!(params.get("clipboard") instanceof ObjectNode clipboard)) throw new IllegalStateException("Legacy paste requires its clipboard");
                    boolean move = params.path("clearSource").asBoolean(false) || clipboard.path("isCut").asBoolean(false);
                    if (move && !params.has("sourceRange")) throw new IllegalStateException("Legacy move paste has no source range");
                    String transfer = move ? "move" : "copy";
                    params.put("transfer", transfer); params.put("clearSource", move); clipboard.put("transfer", transfer); clipboard.remove("isCut");
                    if (!move) params.remove("sourceRange"); changed = true;
                }
            }
            if (value.isContainerNode()) for (JsonNode child : value) pending.push(child);
        }
        return changed;
    }
}
