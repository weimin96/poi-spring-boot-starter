package com.wiblog.poi.exception;

/**
 * @author panwm
 * @since 2023/10/22 12:20
 */
public class ExcelErrorException extends RuntimeException {

    private static final long serialVersionUID = 6769829123439411880L;

    public ExcelErrorException() {
        super();
    }

    public ExcelErrorException(String s) {
        super(s);
    }
}
