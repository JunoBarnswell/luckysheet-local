package com.xc.luckysheet.server.service;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.JsonNode;
import com.xc.luckysheet.server.contract.CopyWorkbookRequest;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.contract.WorkbookLifecycle;
import com.xc.luckysheet.server.contract.WorkbookSnapshotResponse;
import com.xc.luckysheet.server.contract.WorkbookSource;
import com.xc.luckysheet.server.contract.WorkbookStorageLocation;
import com.xc.luckysheet.server.persistence.WorkspaceSpaceEntity;
import com.xc.luckysheet.server.persistence.AuditEntityRepository;
import com.xc.luckysheet.server.persistence.CheckpointEntityRepository;
import com.xc.luckysheet.server.persistence.DataBlockEntityRepository;
import com.xc.luckysheet.server.persistence.OperationEntityRepository;
import com.xc.luckysheet.server.persistence.OutboxEntityRepository;
import com.xc.luckysheet.server.persistence.ShareEntityRepository;
import com.xc.luckysheet.server.persistence.SpaceMemberEntityRepository;
import com.xc.luckysheet.server.persistence.WorkbookAclEntityRepository;
import com.xc.luckysheet.server.persistence.WorkbookEntityRepository;
import com.xc.luckysheet.server.persistence.WorkbookEntity;
import com.xc.luckysheet.server.persistence.WorkbookSourceArtifactEntity;
import com.xc.luckysheet.server.persistence.WorkbookSourceArtifactEntityRepository;
import com.xc.luckysheet.server.persistence.WorkbookUserStateEntityRepository;
import com.xc.luckysheet.server.persistence.WorkspaceFolderEntityRepository;
import com.xc.luckysheet.server.persistence.WorkspaceSpaceEntityRepository;
import org.junit.jupiter.api.Test;

import java.time.Instant;
import java.util.Optional;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.verify;
import static org.mockito.Mockito.when;
import static org.mockito.ArgumentMatchers.eq;

class WorkbookCatalogServiceTest {
    @Test
    void sourceArtifactRequiresEditorAndPersistsVerifiedSha256Bytes() throws Exception {
        WorkbookEntityRepository workbooks = mock(WorkbookEntityRepository.class);
        WorkbookAclEntityRepository acl = mock(WorkbookAclEntityRepository.class);
        WorkbookUserStateEntityRepository states = mock(WorkbookUserStateEntityRepository.class);
        WorkbookSourceArtifactEntityRepository artifacts = mock(WorkbookSourceArtifactEntityRepository.class);
        WorkspaceSpaceEntityRepository spaces = mock(WorkspaceSpaceEntityRepository.class);
        WorkspaceFolderEntityRepository folders = mock(WorkspaceFolderEntityRepository.class);
        SpaceMemberEntityRepository members = mock(SpaceMemberEntityRepository.class);
        WorkspaceService workspace = mock(WorkspaceService.class);
        WorkbookAuthorizationService authorization = mock(WorkbookAuthorizationService.class);
        WorkbookOperationService operations = mock(WorkbookOperationService.class);
        when(authorization.role("book-1", "editor")).thenReturn(Optional.of(com.xc.luckysheet.server.contract.WorkbookAclRole.EDITOR));
        when(workbooks.findById("book-1")).thenReturn(Optional.of(new WorkbookEntity("book-1", "Book", "{}", 0, 0,
                Instant.now(), Instant.now(), "owner", "space-1", null, WorkbookStorageLocation.REMOTE,
                WorkbookSource.NATIVE, WorkbookLifecycle.ACTIVE, null)));
        when(artifacts.findById("book-1")).thenReturn(Optional.empty());
        ObjectMapper mapper = new ObjectMapper().findAndRegisterModules();
        WorkbookCatalogService service = new WorkbookCatalogService(workbooks, acl, states, artifacts, spaces, folders,
                members, workspace, authorization, operations, mock(CheckpointEntityRepository.class),
                mock(OperationEntityRepository.class), mock(OutboxEntityRepository.class), mock(AuditEntityRepository.class),
                mock(ShareEntityRepository.class), mock(DataBlockEntityRepository.class), mapper);

        byte[] content = "xlsx-bytes".getBytes(java.nio.charset.StandardCharsets.UTF_8);
        String checksum = java.util.HexFormat.of().formatHex(java.security.MessageDigest.getInstance("SHA-256").digest(content));
        var response = service.putArtifact("book-1", "report.xlsx", null, checksum, content, "editor");

        assertEquals(checksum, response.checksum());
        assertEquals(content.length, response.byteLength());
        verify(artifacts).save(any(WorkbookSourceArtifactEntity.class));
    }

    @Test
    void copyRewritesWorkbookAndPrintDocumentIdentitiesBeforePersistence() throws Exception {
        WorkbookEntityRepository workbooks = mock(WorkbookEntityRepository.class);
        WorkbookAclEntityRepository acl = mock(WorkbookAclEntityRepository.class);
        WorkbookUserStateEntityRepository states = mock(WorkbookUserStateEntityRepository.class);
        WorkbookSourceArtifactEntityRepository artifacts = mock(WorkbookSourceArtifactEntityRepository.class);
        WorkspaceSpaceEntityRepository spaces = mock(WorkspaceSpaceEntityRepository.class);
        WorkspaceFolderEntityRepository folders = mock(WorkspaceFolderEntityRepository.class);
        SpaceMemberEntityRepository members = mock(SpaceMemberEntityRepository.class);
        WorkspaceService workspace = mock(WorkspaceService.class);
        WorkbookAuthorizationService authorization = mock(WorkbookAuthorizationService.class);
        WorkbookOperationService operations = mock(WorkbookOperationService.class);
        ObjectMapper mapper = new ObjectMapper().findAndRegisterModules();
        Instant now = Instant.parse("2026-08-24T00:00:00Z");
        WorkspaceSpaceEntity space = new WorkspaceSpaceEntity("space-1", "个人空间",
                com.xc.luckysheet.server.contract.WorkspaceSpaceType.PERSONAL, "actor", now, now);
        WorkbookEntity source = new WorkbookEntity("source-1", "Source", "{}", 0, 0, now, now,
                "actor", "space-1", null, WorkbookStorageLocation.REMOTE, WorkbookSource.NATIVE,
                WorkbookLifecycle.ACTIVE, null);
        JsonNode snapshot = mapper.readTree("""
                {"schema":"WorkbookSnapshot","version":10,"unitId":"source-1","name":"Source","dimensionMetrics":{"normalFontFamily":"Calibri","normalFontSizePx":14.6666666667,"maximumDigitWidthPx":7},"calculationSettings":{"mode":"automatic","iterativeCalculation":false,"maximumIterations":100,"maximumChange":0.001,"precisionAsDisplayed":false,"calculateBeforeSave":true,"fullCalculationOnLoad":false},"editingOptions":{"allowEditDirectly":true,"moveAfterEnter":true,"enterDirection":"down","formulaAutoComplete":true,"valueAutoComplete":true,"fixedDecimalPlaces":null},"definedNameModels":[],"dataModel":{"sources":[],"tables":[],"relationships":[],"views":[]},"printDocuments":[{"schema":"PrintDocument","unitId":"source-1","sheetId":"sheet-1"}],"sheets":[{"kind":"worksheet","id":"sheet-1","name":"Sheet1","rowCount":1000,"columnCount":26,"cells":{},"merges":[],"pane":{"kind":"none"},"defaultRowHeightPx":20,"defaultColumnWidthPx":64,"pivots":[],"sparklines":[],"drawings":[],"drawingPayloads":{},"review":{"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}}}]}
                """);
        when(workbooks.findById("source-1")).thenReturn(Optional.of(source));
        when(workbooks.existsById(any())).thenReturn(false);
        when(authorization.role(any(), eq("actor"))).thenReturn(Optional.of(WorkbookAclRole.OWNER));
        when(authorization.role("source-1", "actor")).thenReturn(Optional.of(WorkbookAclRole.VIEWER));
        when(operations.readSnapshot("source-1", "actor")).thenReturn(new WorkbookSnapshotResponse("source-1", snapshot, 0, "checksum"));
        when(workspace.require("space-1", "actor", WorkbookAclRole.EDITOR)).thenReturn(space);
        when(spaces.findById("space-1")).thenReturn(Optional.of(space));
        when(folders.findBySpaceIdOrderByName("space-1")).thenReturn(java.util.List.of());
        when(artifacts.findById(any())).thenReturn(Optional.empty());

        WorkbookCatalogService service = new WorkbookCatalogService(workbooks, acl, states, artifacts, spaces, folders,
                members, workspace, authorization, operations, mock(CheckpointEntityRepository.class),
                mock(OperationEntityRepository.class), mock(OutboxEntityRepository.class), mock(AuditEntityRepository.class),
                mock(ShareEntityRepository.class), mock(DataBlockEntityRepository.class), mapper);

        service.copy("source-1", new CopyWorkbookRequest("Copied", null, null), "actor");

        var capture = org.mockito.ArgumentCaptor.forClass(com.xc.luckysheet.server.persistence.WorkbookEntity.class);
        verify(workbooks).save(capture.capture());
        JsonNode copied = mapper.readTree(capture.getValue().getSnapshotJson());
        assertEquals(capture.getValue().getUnitId(), copied.path("unitId").asText());
        assertEquals("Copied", copied.path("name").asText());
        assertEquals(capture.getValue().getUnitId(), copied.path("printDocuments").get(0).path("unitId").asText());
        assertEquals("source-1", snapshot.path("unitId").asText());
        assertEquals("source-1", snapshot.path("printDocuments").get(0).path("unitId").asText());
    }
}
