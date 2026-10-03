package com.xc.luckysheet.server.service;

import com.xc.luckysheet.server.persistence.DataBlockEntityRepository;
import org.springframework.stereotype.Component;
import org.springframework.scheduling.annotation.Scheduled;

@Component
public class DataBlockGarbageCollectionSchedule {
    private final DataBlockEntityRepository blocks;
    private final DataBlockGarbageCollector collector;
    private static final org.slf4j.Logger LOG = org.slf4j.LoggerFactory.getLogger(DataBlockGarbageCollectionSchedule.class);
    public DataBlockGarbageCollectionSchedule(DataBlockEntityRepository blocks, DataBlockGarbageCollector collector) { this.blocks = blocks; this.collector = collector; }
    private int page = 0;
    @Scheduled(fixedDelay = 60000) public void collectExpiredStaging() {
        var candidates = blocks.expiredMetadata(java.time.Instant.now().minusSeconds(86400), org.springframework.data.domain.PageRequest.of(page, 128));
        page = candidates.size() < 128 ? 0 : page + 1;
        for (var block : candidates) {
            try { collector.collect(block); }
            catch (ServiceException error) { if (!"DATA_BLOCK_REFERENCED".equals(error.code())) LOG.warn("Staging cleanup rejected: {}", error.code()); }
        }
    }
}
