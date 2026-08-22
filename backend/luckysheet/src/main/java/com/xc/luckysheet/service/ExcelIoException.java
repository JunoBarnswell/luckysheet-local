package com.xc.luckysheet.service;

/**
 * Excel 导入/导出的客户端可处理错误（对应 HTTP 400）。
 */
public class ExcelIoException extends RuntimeException {

    public ExcelIoException(String message) {
        super(message);
    }

    public ExcelIoException(String message, Throwable cause) {
        super(message, cause);
    }
}
