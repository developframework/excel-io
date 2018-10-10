package com.github.developframework.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

/**
 * @author qiushui on 2018-10-09.
 * @since 0.1
 */
public final class ExcelIO {

    /**
     * 输出到流
     *
     * @param excelType
     * @param os
     * @return
     */
    public static final ExcelWriter writer(ExcelType excelType, OutputStream os) {
        Workbook workbook;
        switch (excelType) {
            case XLS:
                workbook = new HSSFWorkbook();
                break;
            case XLSX:
                workbook = new XSSFWorkbook();
                break;
            default:
                throw new IllegalArgumentException();
        }
        return new ExcelWriter(workbook, os);
    }

    /**
     * 输出到文件
     *
     * @param excelType
     * @param filename
     * @return
     */
    public static final ExcelWriter writer(ExcelType excelType, String filename) {
        try {
            return writer(excelType, new FileOutputStream(filename));
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 从流读取
     *
     * @param excelType
     * @param inputStream
     * @return
     */
    public static final ExcelReader reader(ExcelType excelType, InputStream inputStream) {
        Workbook workbook;
        try {
            switch (excelType) {
                case XLS:
                    workbook = new HSSFWorkbook(inputStream);
                    break;
                case XLSX:
                    workbook = new XSSFWorkbook(inputStream);
                    break;
                default:
                    throw new IllegalArgumentException();
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        return new ExcelReader(workbook);
    }

    /**
     * 从文件读取
     *
     * @param excelType
     * @param filename
     * @return
     */
    public static final ExcelReader reader(ExcelType excelType, String filename) {
        try {
            return reader(excelType, new FileInputStream(filename));
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        }
    }

}
