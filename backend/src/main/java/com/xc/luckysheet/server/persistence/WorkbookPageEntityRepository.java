package com.xc.luckysheet.server.persistence;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Modifying;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;
import java.time.Instant;
import java.util.Optional;

public interface WorkbookPageEntityRepository extends JpaRepository<WorkbookPageEntity, String> {
    Optional<WorkbookPageEntity> findByUnitIdAndChecksum(String unitId, String checksum);
    @Query("select p.checksum as checksum, p.byteLength as byteLength, p.payloadBase64 as payloadBase64 from WorkbookPageEntity p where p.unitId = :unitId and p.checksum = :checksum")
    Optional<PageContent> findContent(@Param("unitId") String unitId, @Param("checksum") String checksum);
    boolean existsByUnitIdAndChecksum(String unitId, String checksum);
    @Modifying
    @Query(value = "insert into workbook_pages(page_id,unit_id,checksum,byte_length,payload_base64,created_at) values(:id,:unitId,:checksum,:byteLength,:payload,:createdAt)", nativeQuery = true)
    void insertContent(@Param("id") String id, @Param("unitId") String unitId, @Param("checksum") String checksum,
                       @Param("byteLength") long byteLength, @Param("payload") String payload, @Param("createdAt") Instant createdAt);
    void deleteByUnitId(String unitId);
    interface PageContent {
        String getChecksum();
        long getByteLength();
        String getPayloadBase64();
    }
}
