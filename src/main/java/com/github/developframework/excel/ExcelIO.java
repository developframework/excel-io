package com.github.developframework.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 * @author qiushui on 2019-05-18.
 */
public final class ExcelIO {

    /**
     * 写出器
     *
     * @param excelType
     * @return
     */
    public static ExcelWriter writer(ExcelType excelType) {
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
        return new ExcelWriter(workbook);
    }

    /**
     * 从流读取
     *
     * @param excelType
     * @param inputStream
     * @return
     */
    public static ExcelReader reader(ExcelType excelType, InputStream inputStream) {
        Workbook workbook;
        try (inputStream) {
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
     * @param filename
     * @return
     */
    public static ExcelReader reader(String filename) {
        try {
            return reader(ExcelType.parse(filename), new FileInputStream(filename));
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
