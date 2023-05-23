package com.github.developframework.excel.styles;

import lombok.RequiredArgsConstructor;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

/**
 * @author qiushui on 2023-05-23.
 */
@RequiredArgsConstructor
public class CompositeColor {

    private final String color;

    @Override
    public String toString() {
        return color;
    }

    protected void configureFontColor(Workbook workbook, Font font) {
        if (font instanceof XSSFFont) {
            if (color.startsWith("#")) {
                final byte[] rgb = parseRGB(color);
                ((XSSFFont) font).setColor(new XSSFColor(rgb, new DefaultIndexedColorMap()));
            } else {
                font.setColor(IndexedColors.valueOf(color).index);
            }
        } else {
            final HSSFPalette palette = ((HSSFWorkbook) workbook).getCustomPalette();
            HSSFColor hssfColor;
            if (color.startsWith("#")) {
                final byte[] rgb = parseRGB(color);
                hssfColor = palette.findSimilarColor(rgb[0], rgb[1], rgb[2]);
                if (hssfColor == null) {
                    palette.setColorAtIndex((short) 255, rgb[0], rgb[1], rgb[2]);
                    hssfColor = palette.getColor(255);
                }
            } else {
                hssfColor = HSSFColor.HSSFColorPredefined.valueOf(this.color).getColor();
            }
            font.setColor(hssfColor.getIndex());
        }
    }

    protected void configureForegroundColor(Workbook workbook, CellStyle cellStyle) {
        if (cellStyle instanceof XSSFCellStyle) {
            if (color.startsWith("#")) {
                final byte[] rgb = parseRGB(color);
                ((XSSFCellStyle) cellStyle).setFillForegroundColor(new XSSFColor(rgb, null));
            } else {
                cellStyle.setFillForegroundColor(IndexedColors.valueOf(color).index);
            }
        } else {
            final HSSFPalette palette = ((HSSFWorkbook) workbook).getCustomPalette();
            HSSFColor hssfColor;
            if (color.startsWith("#")) {
                final byte[] rgb = parseRGB(color);
                hssfColor = palette.findSimilarColor(rgb[0], rgb[1], rgb[2]);
                if (hssfColor == null) {
                    palette.setColorAtIndex((short) 255, rgb[0], rgb[1], rgb[2]);
                    hssfColor = palette.getColor(255);
                }
            } else {
                hssfColor = HSSFColor.HSSFColorPredefined.valueOf(this.color).getColor();
            }
            cellStyle.setFillForegroundColor(hssfColor.getIndex());
        }
    }

    protected void configureBorderColor(Workbook workbook, CellStyle cellStyle) {
        if (cellStyle instanceof XSSFCellStyle) {
            if (color.startsWith("#")) {
                final byte[] rgb = parseRGB(color);
                final XSSFColor color = new XSSFColor(rgb, null);
                ((XSSFCellStyle) cellStyle).setTopBorderColor(color);
                ((XSSFCellStyle) cellStyle).setRightBorderColor(color);
                ((XSSFCellStyle) cellStyle).setBottomBorderColor(color);
                ((XSSFCellStyle) cellStyle).setLeftBorderColor(color);
            } else {
                cellStyle.setFillForegroundColor(IndexedColors.valueOf(color).index);
            }
        } else {
            final HSSFPalette palette = ((HSSFWorkbook) workbook).getCustomPalette();
            HSSFColor hssfColor;
            if (color.startsWith("#")) {
                final byte[] rgb = parseRGB(color);
                hssfColor = palette.findSimilarColor(rgb[0], rgb[1], rgb[2]);
                if (hssfColor == null) {
                    palette.setColorAtIndex((short) 255, rgb[0], rgb[1], rgb[2]);
                    hssfColor = palette.getColor(255);
                }
            } else {
                hssfColor = HSSFColor.HSSFColorPredefined.valueOf(this.color).getColor();
            }
            cellStyle.setTopBorderColor(hssfColor.getIndex());
            cellStyle.setRightBorderColor(hssfColor.getIndex());
            cellStyle.setBottomBorderColor(hssfColor.getIndex());
            cellStyle.setLeftBorderColor(hssfColor.getIndex());
        }
    }

    private byte[] parseRGB(String rgbStr) {
        final int rgb = Integer.valueOf(rgbStr.substring(1), 16);
        byte r = (byte) (rgb >> 16);
        byte g = (byte) ((rgb & 0x00ff00) >> 8);
        byte b = (byte) (rgb & 0x0000ff);
        return new byte[]{r, g, b};
    }
}
