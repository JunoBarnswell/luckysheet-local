package com.xc.luckysheet.service;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Assert;
import org.junit.Test;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

public class ExcelIoServiceTest {

    private final ExcelIoService service = new ExcelIoService();

    @Test
    public void exportRejectsEmptySheets() {
        try {
            service.exportToXlsx(new JSONObject());
            Assert.fail("expected ExcelIoException");
        } catch (ExcelIoException e) {
            Assert.assertEquals("No sheet data to export", e.getMessage());
        }
    }

    @Test
    public void importRejectsNonXlsx() {
        try {
            service.importFromXlsx(new ByteArrayInputStream(new byte[0]), "book.xls");
            Assert.fail("expected ExcelIoException");
        } catch (ExcelIoException e) {
            Assert.assertTrue(e.getMessage().contains("xlsx"));
        }
    }

    @Test
    public void filterOrderSelectsSingleSheet() {
        JSONObject payload = new JSONObject();
        JSONArray data = new JSONArray();
        JSONObject first = new JSONObject();
        first.put("name", "A");
        JSONObject second = new JSONObject();
        second.put("name", "B");
        data.add(first);
        data.add(second);
        payload.put("data", data);
        JSONObject exportXlsx = new JSONObject();
        exportXlsx.put("order", 1);
        payload.put("exportXlsx", exportXlsx);

        Assert.assertEquals(1, service.filterSheetsByOrder(payload).size());
        Assert.assertEquals("B", service.filterSheetsByOrder(payload).get(0).getString("name"));
    }

    @Test
    public void importAndExportRoundTrip() throws Exception {
        File xlsx = File.createTempFile("roundtrip-", ".xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook();
        workbook.createSheet("Demo").createRow(0).createCell(0).setCellValue("hello");
        FileOutputStream output = new FileOutputStream(xlsx);
        workbook.write(output);
        output.close();
        workbook.close();

        FileInputStream input = new FileInputStream(xlsx);
        JSONObject imported = service.importFromXlsx(input, "demo.xlsx");
        input.close();
        Assert.assertNotNull(imported.getJSONObject("info"));
        Assert.assertTrue(imported.getJSONArray("sheets").size() > 0);

        JSONObject payload = new JSONObject();
        payload.put("title", "demo");
        payload.put("data", imported.getJSONArray("sheets"));
        byte[] exported = service.exportToXlsx(payload);
        Assert.assertTrue(exported.length > 4);
        Assert.assertEquals('P', exported[0]);
        Assert.assertEquals('K', exported[1]);
        Assert.assertTrue(xlsx.delete() || !xlsx.exists());
    }
}
