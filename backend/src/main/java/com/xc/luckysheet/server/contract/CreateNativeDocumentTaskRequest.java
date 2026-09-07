package com.xc.luckysheet.server.contract;

import jakarta.validation.constraints.Max;
import jakarta.validation.constraints.Min;
import jakarta.validation.constraints.NotBlank;
import jakarta.validation.constraints.Pattern;
import jakarta.validation.constraints.Size;

public record CreateNativeDocumentTaskRequest(
        @NotBlank @Size(max = 500) String fileName,
        @Size(max = 500) String name,
        @Size(max = 200) String spaceId,
        @Size(max = 200) String folderId,
        @Min(1) @Max(1073741824L) long byteLength,
        @NotBlank @Pattern(regexp = "[0-9a-fA-F]{64}") String sha256
) { }
