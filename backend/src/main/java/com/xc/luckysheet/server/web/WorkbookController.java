package com.xc.luckysheet.server.web;

import com.xc.luckysheet.server.contract.AclEntry;
import com.xc.luckysheet.server.contract.AclUpdateRequest;
import com.xc.luckysheet.server.contract.AuditRecord;
import com.xc.luckysheet.server.contract.CheckpointResponse;
import com.xc.luckysheet.server.contract.CreateWorkbookRequest;
import com.xc.luckysheet.server.contract.DataBlockMetadata;
import com.xc.luckysheet.server.contract.OperationEnvelope;
import com.xc.luckysheet.server.contract.RestoreRequest;
import com.xc.luckysheet.server.contract.RevisionRecord;
import com.xc.luckysheet.server.contract.WorkbookSnapshotResponse;
import com.xc.luckysheet.server.contract.WorkbookSummary;
import com.xc.luckysheet.server.contract.WorkbookAccessProjection;
import com.xc.luckysheet.server.contract.QueryExecutionRequest;
import com.xc.luckysheet.server.contract.QueryExecutionResponse;
import com.xc.luckysheet.server.contract.ShareCreateRequest;
import com.xc.luckysheet.server.contract.ShareResponse;
import com.xc.luckysheet.server.coordination.WebSocketSessionRegistry;
import com.xc.luckysheet.server.service.ActorIdentity;
import com.xc.luckysheet.server.service.WorkbookOperationService;
import com.xc.luckysheet.server.service.WorkbookDataBlockService;
import com.xc.luckysheet.server.service.QueryExecutionService;
import com.xc.luckysheet.server.service.GuestShareService;
import jakarta.validation.Valid;
import org.springframework.http.ResponseEntity;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.web.bind.annotation.RequestHeader;
import org.springframework.security.core.Authentication;
import org.springframework.web.bind.annotation.DeleteMapping;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.PutMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import java.net.URI;
import java.util.List;

@RestController
@RequestMapping("/api/workbooks")
public class WorkbookController {
    private final WorkbookOperationService operations;
    private final WebSocketSessionRegistry sessions;
    private final QueryExecutionService queries;
    private final GuestShareService guestShares;
    private final WorkbookDataBlockService dataBlocks;

    public WorkbookController(
            WorkbookOperationService operations,
            WebSocketSessionRegistry sessions,
            QueryExecutionService queries,
            GuestShareService guestShares,
            WorkbookDataBlockService dataBlocks
    ) {
        this.operations = operations;
        this.sessions = sessions;
        this.queries = queries;
        this.guestShares = guestShares;
        this.dataBlocks = dataBlocks;
    }

    @PostMapping
    public ResponseEntity<WorkbookSnapshotResponse> create(@Valid @RequestBody CreateWorkbookRequest request, Authentication authentication) {
        ActorIdentity.requireRegisteredActor(authentication);
        WorkbookSnapshotResponse response = operations.create(request, ActorIdentity.subject(authentication));
        return ResponseEntity.created(URI.create("/api/workbooks/" + request.unitId())).body(response);
    }

    @GetMapping
    public List<WorkbookSummary> list(Authentication authentication) {
        return operations.list(ActorIdentity.subject(authentication));
    }

    @GetMapping("/{unitId}/snapshot")
    public WorkbookSnapshotResponse snapshot(@PathVariable String unitId, Authentication authentication) {
        return operations.readSnapshot(unitId, ActorIdentity.subject(authentication));
    }

    @PostMapping("/{unitId}/operations")
    public ResponseEntity<CommitResponse> commit(@PathVariable String unitId, @Valid @RequestBody OperationEnvelope operation, Authentication authentication) {
        WorkbookOperationService.CommitResult result = operations.commit(unitId, operation, ActorIdentity.subject(authentication));
        sessions.broadcastRevision(result.operation());
        return ResponseEntity.status(result.committed() ? 201 : 200).body(new CommitResponse(result.operation()));
    }

    @GetMapping("/{unitId}/revisions")
    public List<RevisionRecord> revisions(@PathVariable String unitId, Authentication authentication) {
        return operations.revisions(unitId, ActorIdentity.subject(authentication));
    }

    @GetMapping("/{unitId}/revisions/{revision}/snapshot")
    public WorkbookSnapshotResponse revisionSnapshot(@PathVariable String unitId, @PathVariable long revision, Authentication authentication) {
        return operations.readRevision(unitId, revision, ActorIdentity.subject(authentication));
    }

    @PostMapping("/{unitId}/checkpoints")
    public CheckpointResponse checkpoint(@PathVariable String unitId, Authentication authentication) {
        return operations.checkpoint(unitId, ActorIdentity.subject(authentication));
    }

    @PostMapping("/{unitId}/restore")
    public WorkbookOperationService.RestoreResult restore(@PathVariable String unitId, @Valid @RequestBody RestoreRequest request, Authentication authentication) {
        WorkbookOperationService.RestoreResult result = operations.restore(unitId, request, ActorIdentity.subject(authentication));
        sessions.broadcastRevision(result.operation());
        return result;
    }

    @GetMapping("/{unitId}/acl")
    public List<AclEntry> acl(@PathVariable String unitId, Authentication authentication) {
        return operations.acl(unitId, ActorIdentity.subject(authentication));
    }

    @GetMapping("/{unitId}/access")
    public WorkbookAccessProjection access(@PathVariable String unitId, Authentication authentication) {
        return operations.accessProjection(unitId, ActorIdentity.subject(authentication));
    }

    @PutMapping("/{unitId}/acl/{subject}")
    public AclEntry grant(@PathVariable String unitId, @PathVariable String subject, @Valid @RequestBody AclUpdateRequest request, Authentication authentication) {
        return operations.grantAcl(unitId, ActorIdentity.subject(authentication), subject, request.role());
    }

    @DeleteMapping("/{unitId}/acl/{subject}")
    public ResponseEntity<Void> revoke(@PathVariable String unitId, @PathVariable String subject, Authentication authentication) {
        operations.revokeAcl(unitId, ActorIdentity.subject(authentication), subject);
        return ResponseEntity.noContent().build();
    }

    @GetMapping("/{unitId}/audit")
    public List<AuditRecord> audit(@PathVariable String unitId, @RequestParam(defaultValue = "100") int limit, Authentication authentication) {
        return operations.audit(unitId, ActorIdentity.subject(authentication), limit);
    }

    @PostMapping("/{unitId}/queries/execute")
    public QueryExecutionResponse executeQuery(
            @PathVariable String unitId,
            @Valid @RequestBody QueryExecutionRequest request,
            Authentication authentication
    ) {
        return queries.execute(unitId, request, ActorIdentity.subject(authentication));
    }

    @PostMapping("/{unitId}/queries/{queryId}/cancel")
    public ResponseEntity<Void> cancelQuery(@PathVariable String unitId, @PathVariable String queryId, Authentication authentication) {
        queries.cancel(unitId, queryId, ActorIdentity.subject(authentication));
        return ResponseEntity.accepted().build();
    }

    @PostMapping("/{unitId}/shares")
    public ShareResponse createShare(
            @PathVariable String unitId,
            @Valid @RequestBody ShareCreateRequest request,
            Authentication authentication
    ) {
        return guestShares.create(unitId, request, ActorIdentity.subject(authentication));
    }

    @GetMapping("/{unitId}/shares")
    public List<ShareResponse> shares(@PathVariable String unitId, Authentication authentication) {
        return guestShares.list(unitId, ActorIdentity.subject(authentication));
    }

    @DeleteMapping("/{unitId}/shares/{shareId}")
    public ResponseEntity<Void> revokeShare(@PathVariable String unitId, @PathVariable java.util.UUID shareId, Authentication authentication) {
        guestShares.revoke(unitId, shareId, ActorIdentity.subject(authentication));
        return ResponseEntity.noContent().build();
    }

    @PutMapping(value = "/{unitId}/data-sources/{sourceId}/blocks/{blockId}", consumes = MediaType.APPLICATION_OCTET_STREAM_VALUE)
    public DataBlockMetadata putDataBlock(
            @PathVariable String unitId,
            @PathVariable String sourceId,
            @PathVariable String blockId,
            @RequestHeader("X-Content-SHA256") String checksum,
            @RequestBody byte[] content,
            Authentication authentication
    ) {
        return dataBlocks.put(unitId, sourceId, blockId, checksum, content, ActorIdentity.subject(authentication));
    }

    @GetMapping(value = "/{unitId}/data-sources/{sourceId}/blocks/{blockId}", produces = MediaType.APPLICATION_OCTET_STREAM_VALUE)
    public ResponseEntity<byte[]> getDataBlock(
            @PathVariable String unitId,
            @PathVariable String sourceId,
            @PathVariable String blockId,
            Authentication authentication
    ) {
        var block = dataBlocks.get(unitId, sourceId, blockId, ActorIdentity.subject(authentication));
        return ResponseEntity.ok()
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .contentLength(block.byteLength())
                .header("X-Content-SHA256", block.checksum())
                .header(HttpHeaders.CACHE_CONTROL, "private, max-age=0")
                .body(block.content());
    }

    @DeleteMapping("/{unitId}/data-sources/{sourceId}/blocks/{blockId}")
    public ResponseEntity<Void> deleteDataBlock(
            @PathVariable String unitId,
            @PathVariable String sourceId,
            @PathVariable String blockId,
            Authentication authentication
    ) {
        dataBlocks.delete(unitId, sourceId, blockId, ActorIdentity.subject(authentication));
        return ResponseEntity.noContent().build();
    }
}
