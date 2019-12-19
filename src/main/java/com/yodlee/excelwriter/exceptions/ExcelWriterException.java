package com.yodlee.excelwriter.exceptions;

public class ExcelWriterException extends Exception {
    public ExcelWriterException(){
    }
    public ExcelWriterException(String message) {
        super(message);
    }

    public ExcelWriterException(String message, Throwable cause) {
        super(message, cause);
    }

    public ExcelWriterException(Throwable cause) {
        super(cause);
    }

    public ExcelWriterException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }

}
