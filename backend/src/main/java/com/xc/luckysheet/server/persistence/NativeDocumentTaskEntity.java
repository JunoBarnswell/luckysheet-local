package com.xc.luckysheet.server.persistence;

import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;
import org.hibernate.annotations.JdbcTypeCode;
import org.hibernate.type.SqlTypes;
import java.time.Instant;

/** Durable actor-owned upload and publication state. No client path is persisted or accepted. */
@Entity
@Table(name = "native_document_task")
public class NativeDocumentTaskEntity {
    @Id @Column(name = "task_id", length = 36, nullable = false)
    private String taskId;
    @Column(name = "actor_subject", length = 500, nullable = false)
    private String actorSubject;
    @Column(name = "file_name", length = 500, nullable = false)
    private String fileName;
    @Column(name = "workbook_name", length = 500)
    private String name;
    @Column(name = "space_id", length = 200)
    private String spaceId;
    @Column(name = "folder_id", length = 200)
    private String folderId;
    @Column(name = "byte_length", nullable = false)
    private long byteLength;
    @Column(name = "sha256", length = 64)
    private String sha256;
    @Column(name = "uploaded_bytes", nullable = false)
    private long uploadedBytes;
    @Column(name = "state", length = 16, nullable = false)
    private String state;
    @JdbcTypeCode(SqlTypes.LONGVARCHAR) @Column(name = "result_json")
    private String resultJson;
    @Column(name = "error_code", length = 100)
    private String errorCode;
    @JdbcTypeCode(SqlTypes.LONGVARCHAR) @Column(name = "error_message")
    private String errorMessage;
    @Column(name = "created_at", nullable = false)
    private Instant createdAt;
    @Column(name = "updated_at", nullable = false)
    private Instant updatedAt;

    protected NativeDocumentTaskEntity() { }
    public NativeDocumentTaskEntity(String id, String actor, String fileName, String name, String spaceId,
            String folderId, long byteLength, String sha256) {
        this.taskId = id; this.actorSubject = actor; this.fileName = fileName; this.name = name;
        this.spaceId = spaceId; this.folderId = folderId; this.byteLength = byteLength; this.sha256 = sha256;
        state = "uploading"; createdAt = Instant.now(); updatedAt = createdAt;
    }
    public String getTaskId() { return taskId; }
    public String getActorSubject() { return actorSubject; }
    public String getFileName() { return fileName; }
    public String getName() { return name; }
    public String getSpaceId() { return spaceId; }
    public String getFolderId() { return folderId; }
    public long getByteLength() { return byteLength; }
    public String getSha256() { return sha256; }
    public long getUploadedBytes() { return uploadedBytes; }
    public String getState() { return state; }
    public String getResultJson() { return resultJson; }
    public String getErrorCode() { return errorCode; }
    public String getErrorMessage() { return errorMessage; }
    public void uploaded(long bytes) { uploadedBytes = bytes; updatedAt = Instant.now(); }
    public void importing() { state = "importing"; updatedAt = Instant.now(); }
    public void completed(String result) { state = "completed"; resultJson = result; updatedAt = Instant.now(); }
    public void cancelled() { state = "cancelled"; updatedAt = Instant.now(); }
    public void failed(String code, String message) {
        state = "failed"; errorCode = code; errorMessage = message; updatedAt = Instant.now();
    }
}
