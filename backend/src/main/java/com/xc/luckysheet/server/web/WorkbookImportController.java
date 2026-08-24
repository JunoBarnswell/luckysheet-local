package com.xc.luckysheet.server.web;

import com.xc.luckysheet.server.contract.WorkbookImportResponse;
import com.xc.luckysheet.server.service.ActorIdentity;
import com.xc.luckysheet.server.service.WorkbookCatalogService;
import jakarta.validation.constraints.Size;
import org.springframework.http.MediaType;
import org.springframework.security.core.Authentication;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RequestPart;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

@RestController
@RequestMapping("/api/workbook-imports")
public class WorkbookImportController {
    private final WorkbookCatalogService catalog;

    public WorkbookImportController(WorkbookCatalogService catalog) {
        this.catalog = catalog;
    }

    @PostMapping(consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public WorkbookImportResponse importWorkbook(
            @RequestPart("file") MultipartFile file,
            @RequestPart("snapshot") String snapshot,
            @RequestParam int xlsxCodecVersion,
            @RequestParam String detectedFeatures,
            @RequestParam String capabilityReport,
            @RequestParam(required = false) @Size(max = 500) String name,
            @RequestParam(required = false) String spaceId,
            @RequestParam(required = false) String folderId,
            Authentication authentication
    ) {
        ActorIdentity.requireRegisteredActor(authentication);
        return catalog.importXlsx(file, name, spaceId, folderId, snapshot, xlsxCodecVersion, detectedFeatures, capabilityReport, ActorIdentity.subject(authentication));
    }
}
