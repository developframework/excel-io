package com.github.developframework.excel.styles;

import lombok.Getter;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;
import java.lang.reflect.Method;
import java.util.LinkedList;
import java.util.List;

/**
 * @author qiushui on 2023-05-19.
 */
public class CellStyleKey {

    private static final Class<?>[] ITEM_KEY_CLASSES = new Class[]{
            FontKey.class,
            AlignmentKey.class,
            ForegroundKey.class,
            BorderKey.class,
            DataFormatKey.class
    };

    private final List<ItemKey> itemKeys = new LinkedList<>();

    public static CellStyleKey parse(String key) {
        final CellStyleKey cellStyleKey = new CellStyleKey();
        final String[] items = key.split(";\\s*");
        for (String item : items) {
            Class<?> matcheClass = null;
            x:
            for (Class<?> itemKeyClass : ITEM_KEY_CLASSES) {
                final String[] prefixArray = itemKeyClass.getAnnotation(Prefix.class).value();
                for (String prefix : prefixArray) {
                    if (item.startsWith(prefix)) {
                        item = item.substring(prefix.length());
                        matcheClass = itemKeyClass;
                        break x;
                    }
                }
            }
            if (matcheClass != null) {
                try {
                    final Method method = matcheClass.getMethod("of", String.class);
                    method.setAccessible(true);
                    cellStyleKey.itemKeys.add(
                            (ItemKey) method.invoke(null, item)
                    );
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        }
        return cellStyleKey;
    }

    @FunctionalInterface
    public interface ItemKey {

        void configureCellStyle(Workbook workbook, CellStyle cellStyle);
    }

    @Retention(RetentionPolicy.RUNTIME)
    @Target(ElementType.TYPE)
    public @interface Prefix {

        String[] value();
    }

    @Getter
    @Prefix({"font#", "f#"})
    protected static class FontKey implements ItemKey {

        private short heightInPoints;

        private short color;

        private XSSFColor xssfColor;

        private String fontName;

        private boolean italic;

        private boolean bold;

        @Override
        public void configureCellStyle(Workbook workbook, CellStyle cellStyle) {
            final Font font = workbook.createFont();
            font.setFontName(fontName);
            font.setFontHeightInPoints(heightInPoints);
            font.setItalic(italic);
            font.setBold(bold);

            if (font instanceof XSSFFont) {
                ((XSSFFont) font).setColor(xssfColor);
            } else {
                font.setColor(color);
            }
            cellStyle.setFont(font);
        }

        public static FontKey of(String key) {
            final FontKey fontKey = new FontKey();
            final String[] parts = key.split("-");
            for (String part : parts) {
                if (StringUtils.isNumeric(part)) {
                    fontKey.heightInPoints = Short.parseShort(part);
                } else if (part.equalsIgnoreCase("italic")) {
                    fontKey.italic = true;
                } else if (part.equalsIgnoreCase("bold")) {
                    fontKey.bold = true;
                } else if (part.startsWith("#")) {
                    final int rgb = Integer.valueOf(part.substring(1), 16);
                    int r = rgb >> 16;
                    int g = (rgb & 0x00ff00) >> 8;
                    int b = rgb & 0x0000ff;
                    fontKey.xssfColor = new XSSFColor(new java.awt.Color(r, g, b));
                } else {
                    fontKey.fontName = part;
                }
            }
            return fontKey;
        }
    }

    @Getter
    @Prefix({"align#", "a#"})
    protected static class AlignmentKey implements ItemKey {

        private HorizontalAlignment horizontalAlignment = HorizontalAlignment.CENTER;

        private VerticalAlignment verticalAlignment = VerticalAlignment.CENTER;

        @Override
        public void configureCellStyle(Workbook workbook, CellStyle cellStyle) {
            cellStyle.setAlignment(horizontalAlignment);
            cellStyle.setVerticalAlignment(verticalAlignment);
        }

        public static AlignmentKey of(String key) {
            AlignmentKey alignmentKey = new AlignmentKey();
            final String[] parts = key.split("-");
            for (String part : parts) {
                if (part.startsWith("v:")) {
                    alignmentKey.verticalAlignment = VerticalAlignment.valueOf(part.substring(2).toUpperCase());
                } else if (part.startsWith("h:")) {
                    alignmentKey.horizontalAlignment = HorizontalAlignment.valueOf(part.substring(2).toUpperCase());
                }
            }
            return alignmentKey;
        }
    }

    @Getter
    @Prefix({"foreground#", "fg#"})
    protected static class ForegroundKey implements ItemKey {

        private IndexedColors indexedColor = IndexedColors.AUTOMATIC;

        private XSSFColor xssfColor;

        private FillPatternType fillPatternType = FillPatternType.SOLID_FOREGROUND;

        @Override
        public void configureCellStyle(Workbook workbook, CellStyle cellStyle) {
            if (cellStyle instanceof XSSFCellStyle) {
                XSSFCellStyle xssfCellStyle = (XSSFCellStyle) cellStyle;
                xssfCellStyle.setFillForegroundColor(xssfColor);
            }
            cellStyle.setFillPattern(fillPatternType);
        }

        public static ForegroundKey of(String key) {
            ForegroundKey foregroundKey = new ForegroundKey();
            final String[] parts = key.split("-");
            for (String part : parts) {
                if (part.startsWith("#")) {
                    final int rgb = Integer.valueOf(part.substring(1), 16);
                    int r = rgb >> 16;
                    int g = (rgb & 0x00ff00) >> 8;
                    int b = rgb & 0x0000ff;
                    foregroundKey.xssfColor = new XSSFColor(new java.awt.Color(r, g, b));
                } else {
                    foregroundKey.fillPatternType = FillPatternType.valueOf(part.toUpperCase());
                }
            }
            return foregroundKey;
        }
    }

    @Getter
    @Prefix({"border#", "b#"})
    protected static class BorderKey implements ItemKey {

        private BorderStyle borderStyle;

        private XSSFColor xssfColor;

        @Override
        public void configureCellStyle(Workbook workbook, CellStyle cellStyle) {
            cellStyle.setBorderTop(borderStyle);
            cellStyle.setBorderRight(borderStyle);
            cellStyle.setBorderBottom(borderStyle);
            cellStyle.setBorderLeft(borderStyle);

            if (cellStyle instanceof XSSFCellStyle) {
                XSSFCellStyle xssfCellStyle = (XSSFCellStyle) cellStyle;
                xssfCellStyle.setTopBorderColor(xssfColor);
                xssfCellStyle.setRightBorderColor(xssfColor);
                xssfCellStyle.setBottomBorderColor(xssfColor);
                xssfCellStyle.setLeftBorderColor(xssfColor);
            }
        }

        public static BorderKey of(String key) {
            BorderKey borderKey = new BorderKey();
            final String[] parts = key.split("-");
            for (String part : parts) {
                if (part.startsWith("#")) {
                    final int rgb = Integer.valueOf(part.substring(1), 16);
                    int r = rgb >> 16;
                    int g = (rgb & 0x00ff00) >> 8;
                    int b = rgb & 0x0000ff;
                    borderKey.xssfColor = new XSSFColor(new java.awt.Color(r, g, b));
                } else {
                    borderKey.borderStyle = BorderStyle.valueOf(part.toUpperCase());
                }
            }
            return borderKey;
        }
    }

    @Getter
    @Prefix({"dataFormat#", "df#"})
    protected static class DataFormatKey implements ItemKey {

        private String format;

        @Override
        public void configureCellStyle(Workbook workbook, CellStyle cellStyle) {
            cellStyle.setDataFormat(workbook.createDataFormat().getFormat(format));
        }

        public static DataFormatKey of(String key) {
            DataFormatKey dataFormatKey = new DataFormatKey();
            dataFormatKey.format = key;
            return dataFormatKey;
        }
    }
}
