package com.xc.luckysheet.controller;

import com.alibaba.fastjson.JSONObject;
import com.xc.luckysheet.service.ExcelIoException;
import com.xc.luckysheet.service.ExcelIoService;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.net.URLEncoder;

/**
 * 对齐 luckyexcel-node 的 Excel 导入/导出 HTTP 契约。
 */
@Slf4j
@RestController
@CrossOrigin(origins = "*")
@Api(description = "Excel 导入导出")
public class ExcelIoController {

    @Autowired
    private ExcelIoService excelIoService;

    @ApiOperation(value = "导出 xlsx", notes = "将 Luckysheet toJson() 转为 xlsx 二进制")
    @PostMapping("/luckyToXlsx")
    public void exportXlsx(@RequestBody(required = false) JSONObject body, HttpServletResponse response) throws IOException {
        try {
            byte[] data = excelIoService.exportToXlsx(body);
            String filename = excelIoService.sanitizeFilename(body);
            response.setStatus(HttpServletResponse.SC_OK);
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.setHeader("Content-Disposition", "attachment; filename=\"" + URLEncoder.encode(filename, "UTF-8") + "\"");
            response.setContentLength(data.length);
            response.getOutputStream().write(data);
            response.getOutputStream().flush();
        } catch (ExcelIoException e) {
            writeError(response, HttpServletResponse.SC_BAD_REQUEST, e.getMessage());
        } catch (Exception e) {
            log.error("export xlsx failed", e);
            writeError(response, HttpServletResponse.SC_BAD_REQUEST, e.getMessage() == null ? "Failed to export xlsx" : e.getMessage());
        }
    }

    @ApiOperation(value = "导入 xlsx", notes = "上传 xlsx，返回 { info, sheets }")
    @PostMapping("/luckyexcel/upload")
    public void importUpload(@RequestParam(value = "file", required = false) MultipartFile file,
                             HttpServletResponse response) throws IOException {
        try {
            if (file == null || file.isEmpty()) {
                writeError(response, HttpServletResponse.SC_BAD_REQUEST, "Missing file field \"file\"");
                return;
            }
            JSONObject result = excelIoService.importFromXlsx(file.getInputStream(), file.getOriginalFilename());
            response.setStatus(HttpServletResponse.SC_OK);
            response.setContentType("application/json;charset=UTF-8");
            response.getWriter().write(result.toJSONString());
        } catch (ExcelIoException e) {
            writeError(response, HttpServletResponse.SC_BAD_REQUEST, e.getMessage());
        } catch (Exception e) {
            log.error("import xlsx failed", e);
            writeError(response, HttpServletResponse.SC_BAD_REQUEST, e.getMessage() == null ? "Invalid xlsx file" : e.getMessage());
        }
    }

    private void writeError(HttpServletResponse response, int status, String message) throws IOException {
        response.reset();
        response.setStatus(status);
        response.setContentType("application/json;charset=UTF-8");
        JSONObject error = new JSONObject();
        error.put("error", message);
        response.getWriter().write(error.toJSONString());
    }
}
