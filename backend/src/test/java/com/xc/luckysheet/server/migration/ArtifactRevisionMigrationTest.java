package com.xc.luckysheet.server.migration;

import org.junit.jupiter.api.Test;
import java.nio.charset.StandardCharsets;
import java.sql.DriverManager;
import static org.junit.jupiter.api.Assertions.*;

class ArtifactRevisionMigrationTest {
    @Test
    void existingArtifactBytesRemainIntactAndRevisionRemainsUnproven() throws Exception {
        try (var connection = DriverManager.getConnection("jdbc:h2:mem:artifact-" + java.util.UUID.randomUUID());
             var statement = connection.createStatement()) {
            statement.execute("create table workbook_source_artifact(unit_id varchar primary key, content varbinary, checksum varchar)");
            statement.execute("insert into workbook_source_artifact values('existing', X'01020304', 'original-checksum')");
            try (var resource = getClass().getResourceAsStream("/db/migration/h2/V10__artifact_revision.sql")) {
                assertNotNull(resource);
                statement.execute(new String(resource.readAllBytes(), StandardCharsets.UTF_8));
            }
            try (var rows = statement.executeQuery("select content,checksum,source_revision from workbook_source_artifact where unit_id='existing'")) {
                assertTrue(rows.next());
                assertArrayEquals(new byte[]{1, 2, 3, 4}, rows.getBytes("content"));
                assertEquals("original-checksum", rows.getString("checksum"));
                assertNull(rows.getObject("source_revision"));
            }
        }
    }
}
