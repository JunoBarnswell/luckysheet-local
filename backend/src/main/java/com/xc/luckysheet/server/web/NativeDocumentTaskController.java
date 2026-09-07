package com.xc.luckysheet.server.web;

import com.xc.luckysheet.server.contract.CreateNativeDocumentTaskRequest;
import com.xc.luckysheet.server.contract.NativeDocumentTaskResponse;
import com.xc.luckysheet.server.service.ActorIdentity;
import com.xc.luckysheet.server.service.NativeDocumentTaskService;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.validation.Valid;
import org.springframework.http.MediaType;
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
import java.io.IOException;

@RestController
@RequestMapping("/api/workbook-imports/tasks")
public class NativeDocumentTaskController {
    private final NativeDocumentTaskService tasks;
    public NativeDocumentTaskController(NativeDocumentTaskService tasks) { this.tasks = tasks; }

    @PostMapping(consumes = MediaType.APPLICATION_JSON_VALUE)
    public NativeDocumentTaskResponse create(@Valid @RequestBody CreateNativeDocumentTaskRequest request,
            Authentication authentication) {
        return tasks.create(request, actor(authentication));
    }
    @GetMapping("/{id}")
    public NativeDocumentTaskResponse read(@PathVariable String id, Authentication authentication) {
        return tasks.read(id, actor(authentication));
    }
    @PutMapping(value = "/{id}/chunks", consumes = MediaType.APPLICATION_OCTET_STREAM_VALUE)
    public NativeDocumentTaskResponse upload(@PathVariable String id, @RequestParam long offset,
            HttpServletRequest request, Authentication authentication) throws IOException {
        String actor = actor(authentication);
        return tasks.upload(id, offset, request.getInputStream(), actor);
    }
    @PostMapping("/{id}/commit")
    public NativeDocumentTaskResponse commit(@PathVariable String id, Authentication authentication) {
        return tasks.commit(id, actor(authentication));
    }
    @DeleteMapping("/{id}")
    public NativeDocumentTaskResponse cancel(@PathVariable String id, Authentication authentication) {
        return tasks.cancel(id, actor(authentication));
    }
    private static String actor(Authentication authentication) {
        ActorIdentity.requireRegisteredActor(authentication);
        return ActorIdentity.subject(authentication);
    }
}
