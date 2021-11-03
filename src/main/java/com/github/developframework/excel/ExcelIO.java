package com.github.developframework.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.UncheckedIOException;

/**
 * @author qiushui on 2019-05-18.
 */
@SuppressWarnings("unused")
public final class ExcelIO {

    /**
     * 写出器
     *
     * @param excelType Excel类型
     * @return 写出器
     */
    public static ExcelWriter writer(ExcelType excelType) {
        Workbook workbook;
        switch (excelType) {
            case XLS:
                workbook = new HSSFWorkbook();
                break;
            case XLSX:
                workbook = new SXSSFWorkbook();
                break;
            default:
                throw new IllegalArgumentException();
        }
        return new ExcelWriter(workbook);
    }

    /**
     * 从流读取
     *
     * @param excelType   Excel类型
     * @param inputStream 输入流
     * @return 读取器
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
            throw new UncheckedIOException(e);
        }
        return new ExcelReader(workbook);
    }

    /**
     * 从文件读取
     *
     * @param filename 文件名
     * @return 读取器
     */
    public static ExcelReader reader(String filename) {
        try {
            return reader(ExcelType.parse(filename), new FileInputStream(filename));
        } catch (IOException e) {
            throw new UncheckedIOException(e);
        }
    }

    /**
     * 从流读取(带密码)
     *
     * @param inputStream 输入流
     * @param password    密码
     * @return 读取器
     */
    public static ExcelReader readerWithPassword(InputStream inputStream, String password) {
        try (inputStream) {
            return new ExcelReader(WorkbookFactory.create(inputStream, password));
        } catch (IOException e) {
            throw new UncheckedIOException(e);
        }
    }

    /**
     * 从文件读取(带密码)
     *
     * @param filename 文件名
     * @param password 密码
     * @return 读取器
     */
    public static ExcelReader readerWithPassword(String filename, String password) {
        try {
            return readerWithPassword(new FileInputStream(filename), password);
        } catch (IOException e) {
            throw new UncheckedIOException(e);
        }
    }
}
