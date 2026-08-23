package com.xc.luckysheet.server.web;

import com.xc.luckysheet.server.contract.FolderRequest;
import com.xc.luckysheet.server.contract.FolderResponse;
import com.xc.luckysheet.server.service.ActorIdentity;
import com.xc.luckysheet.server.service.WorkspaceService;
import jakarta.validation.Valid;
import org.springframework.http.ResponseEntity;
import org.springframework.security.core.Authentication;
import org.springframework.web.bind.annotation.DeleteMapping;
import org.springframework.web.bind.annotation.PatchMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("/api/folders")
public class FolderController {
    private final WorkspaceService workspace;

    public FolderController(WorkspaceService workspace) {
        this.workspace = workspace;
    }

    @PatchMapping("/{folderId}")
    public FolderResponse update(@PathVariable String folderId, @Valid @RequestBody FolderRequest request, Authentication authentication) {
        ActorIdentity.requireRegisteredActor(authentication);
        return workspace.updateFolder(folderId, request, ActorIdentity.subject(authentication));
    }

    @DeleteMapping("/{folderId}")
    public ResponseEntity<Void> delete(@PathVariable String folderId, Authentication authentication) {
        ActorIdentity.requireRegisteredActor(authentication);
        workspace.deleteFolder(folderId, ActorIdentity.subject(authentication));
        return ResponseEntity.noContent().build();
    }
}
