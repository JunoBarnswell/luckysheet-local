package com.xc.luckysheet.server.web;

import com.xc.luckysheet.server.contract.CreateSpaceRequest;
import com.xc.luckysheet.server.contract.FolderRequest;
import com.xc.luckysheet.server.contract.FolderResponse;
import com.xc.luckysheet.server.contract.SpaceMemberRequest;
import com.xc.luckysheet.server.contract.SpaceMemberResponse;
import com.xc.luckysheet.server.contract.SpaceResponse;
import com.xc.luckysheet.server.service.ActorIdentity;
import com.xc.luckysheet.server.service.WorkspaceService;
import jakarta.validation.Valid;
import org.springframework.http.ResponseEntity;
import org.springframework.security.core.Authentication;
import org.springframework.web.bind.annotation.DeleteMapping;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.PutMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;

@RestController
@RequestMapping("/api/spaces")
public class WorkspaceController {
    private final WorkspaceService workspace;

    public WorkspaceController(WorkspaceService workspace) {
        this.workspace = workspace;
    }

    @GetMapping
    public List<SpaceResponse> list(Authentication authentication) {
        ActorIdentity.requireRegisteredActor(authentication);
        return workspace.list(ActorIdentity.subject(authentication));
    }

    @PostMapping
    public SpaceResponse create(@Valid @RequestBody CreateSpaceRequest request, Authentication authentication) {
        ActorIdentity.requireRegisteredActor(authentication);
        return workspace.create(request, ActorIdentity.subject(authentication));
    }

    @GetMapping("/{spaceId}/folders")
    public List<FolderResponse> folders(@PathVariable String spaceId, Authentication authentication) {
        ActorIdentity.requireRegisteredActor(authentication);
        return workspace.listFolders(spaceId, ActorIdentity.subject(authentication));
    }

    @PostMapping("/{spaceId}/folders")
    public FolderResponse createFolder(@PathVariable String spaceId, @Valid @RequestBody FolderRequest request, Authentication authentication) {
        ActorIdentity.requireRegisteredActor(authentication);
        return workspace.createFolder(spaceId, request, ActorIdentity.subject(authentication));
    }

    @DeleteMapping("/{spaceId}/folders/{folderId}")
    public ResponseEntity<Void> deleteFolder(@PathVariable String spaceId, @PathVariable String folderId, Authentication authentication) {
        ActorIdentity.requireRegisteredActor(authentication);
        workspace.deleteFolder(spaceId, folderId, ActorIdentity.subject(authentication));
        return ResponseEntity.noContent().build();
    }

    @GetMapping("/{spaceId}/members")
    public List<SpaceMemberResponse> members(@PathVariable String spaceId, Authentication authentication) {
        ActorIdentity.requireRegisteredActor(authentication);
        return workspace.listMembers(spaceId, ActorIdentity.subject(authentication));
    }

    @PutMapping("/{spaceId}/members/{subject}")
    public SpaceMemberResponse updateMember(@PathVariable String spaceId, @PathVariable String subject,
                                            @Valid @RequestBody SpaceMemberRequest request, Authentication authentication) {
        ActorIdentity.requireRegisteredActor(authentication);
        return workspace.upsertMember(spaceId, subject, request, ActorIdentity.subject(authentication));
    }

    @DeleteMapping("/{spaceId}/members/{subject}")
    public ResponseEntity<Void> deleteMember(@PathVariable String spaceId, @PathVariable String subject, Authentication authentication) {
        ActorIdentity.requireRegisteredActor(authentication);
        workspace.removeMember(spaceId, subject, ActorIdentity.subject(authentication));
        return ResponseEntity.noContent().build();
    }
}
