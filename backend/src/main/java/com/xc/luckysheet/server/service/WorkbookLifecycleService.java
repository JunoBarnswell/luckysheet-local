package com.xc.luckysheet.server.service;

import com.xc.luckysheet.server.contract.WorkbookLifecycle;
import com.xc.luckysheet.server.persistence.WorkbookEntity;
import com.xc.luckysheet.server.persistence.WorkbookEntityRepository;
import org.springframework.stereotype.Service;

/** Single lifecycle policy shared by operation-adjacent write services. */
@Service
public class WorkbookLifecycleService {
    private final WorkbookEntityRepository workbooks;

    public WorkbookLifecycleService(WorkbookEntityRepository workbooks) {
        this.workbooks = workbooks;
    }

    public void requireActive(String unitId) {
        WorkbookEntity workbook = workbooks.findById(unitId)
                .orElseThrow(() -> ServiceException.notFound("Workbook not found: " + unitId));
        if (workbook.getLifecycle() != WorkbookLifecycle.ACTIVE) {
            throw ServiceException.trashed("Workbook is in trash");
        }
    }
}
