package com.xc.luckysheet.service;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import io.github.autoffice.luckysheet.LuckysheetConverter;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.List;

/**
 * 将 luckyexcel-node 的导入/导出契约适配到 luckysheet-lib。
 */
@Slf4j
@Service
public class ExcelIoService {

    public byte[] exportToXlsx(JSONObject payload) {
        List<JSONObject> sheets = filterSheetsByOrder(payload);
        if (sheets.isEmpty()) {
            throw new ExcelIoException("No sheet data to export");
        }

        JSONObject workbook = buildWorkbookJson(payload, sheets);
        File jsonFile = null;
        File xlsxFile = null;
        try {
            jsonFile = File.createTempFile("luckysheet-export-", ".json");
            xlsxFile = File.createTempFile("luckysheet-export-", ".xlsx");
            Files.write(jsonFile.toPath(), workbook.toJSONString().getBytes(StandardCharsets.UTF_8));
            LuckysheetConverter.luckysheetToExcel(jsonFile.getAbsolutePath(), xlsxFile.getAbsolutePath());
            return Files.readAllBytes(xlsxFile.toPath());
        } catch (ExcelIoException e) {
            throw e;
        } catch (Exception e) {
            log.error("export xlsx failed", e);
            throw new ExcelIoException(e.getMessage() == null ? "Failed to export xlsx" : e.getMessage(), e);
        } finally {
            deleteQuietly(jsonFile);
            deleteQuietly(xlsxFile);
        }
    }

    public JSONObject importFromXlsx(InputStream input, String filename) {
        if (input == null) {
            throw new ExcelIoException("Missing file field \"file\"");
        }
        if (filename == null || !filename.toLowerCase().endsWith(".xlsx")) {
            throw new ExcelIoException("Only .xlsx files are supported");
        }

        File xlsxFile = null;
        try {
            xlsxFile = File.createTempFile("luckysheet-import-", ".xlsx");
            Files.copy(input, xlsxFile.toPath(), StandardCopyOption.REPLACE_EXISTING);
            String json = LuckysheetConverter.excelToLuckySheetJson(xlsxFile.getAbsolutePath());
            JSONObject result = JSON.parseObject(json);
            if (result == null || result.getJSONArray("sheets") == null) {
                throw new ExcelIoException("Invalid xlsx file");
            }
            return result;
        } catch (ExcelIoException e) {
            throw e;
        } catch (Exception e) {
            log.error("import xlsx failed", e);
            throw new ExcelIoException(e.getMessage() == null ? "Invalid xlsx file" : e.getMessage(), e);
        } finally {
            deleteQuietly(xlsxFile);
        }
    }

    public String sanitizeFilename(JSONObject payload) {
        String title = "luckysheet";
        if (payload != null) {
            String raw = payload.getString("title");
            if (raw != null && raw.trim().length() > 0) {
                title = raw.trim();
            }
        }
        return title.replaceAll("[\\\\/:*?\"<>|]", "_") + ".xlsx";
    }

    List<JSONObject> filterSheetsByOrder(JSONObject payload) {
        List<JSONObject> sheets = new ArrayList<JSONObject>();
        if (payload == null) {
            return sheets;
        }
        JSONArray data = payload.getJSONArray("data");
        if (data != null) {
            for (int i = 0; i < data.size(); i++) {
                JSONObject sheet = data.getJSONObject(i);
                if (sheet != null) {
                    sheets.add(sheet);
                }
            }
        }

        JSONObject exportXlsx = payload.getJSONObject("exportXlsx");
        if (exportXlsx == null) {
            return sheets;
        }
        Object order = exportXlsx.get("order");
        if (order == null || "all".equals(String.valueOf(order))) {
            return sheets;
        }
        try {
            int index = Integer.parseInt(String.valueOf(order));
            if (index >= 0 && index < sheets.size()) {
                List<JSONObject> selected = new ArrayList<JSONObject>();
                selected.add(sheets.get(index));
                return selected;
            }
        } catch (NumberFormatException ignored) {
            // 与 Node 一致：无法解析时导出全部 sheet
        }
        return sheets;
    }

    JSONObject buildWorkbookJson(JSONObject payload, List<JSONObject> sheets) {
        JSONArray cleaned = new JSONArray();
        for (JSONObject sheet : sheets) {
            JSONObject copy = JSON.parseObject(sheet.toJSONString());
            copy.remove("data");
            cleaned.add(copy);
        }

        JSONObject info = new JSONObject();
        info.put("name", sanitizeFilename(payload).replaceAll("\\.xlsx$", ""));

        JSONObject workbook = new JSONObject();
        workbook.put("info", info);
        workbook.put("sheets", cleaned);
        return workbook;
    }

    private void deleteQuietly(File file) {
        if (file != null && file.exists() && !file.delete()) {
            file.deleteOnExit();
        }
    }
}
