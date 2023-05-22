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
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

/**
 * @author qiushui on 2023-05-19.
 */
public class CellStyleKey {

    @SuppressWarnings("unchecked")
    private static final Class<? extends ItemKey>[] ITEM_KEY_CLASSES = new Class[]{
            FontKey.class,
            AlignmentKey.class,
            ForegroundKey.class,
            BorderKey.class,
            DataFormatKey.class
    };

    private final List<ItemKey> itemKeys = new LinkedList<>();

    public static boolean isCellStyleKey(String key) {
        return key.matches("^(\\s*\\w+\\s*\\{(.+?)*}\\s*)+$");
    }

    public static CellStyleKey parse(String key) {
        final CellStyleKey cellStyleKey = new CellStyleKey();
        final Map<Class<? extends ItemKey>, Map<String, String>> itemKeyMetadataMap = disassembleItemKey(key);

        itemKeyMetadataMap.forEach((itemKeyClass, properties) -> {
            try {
                cellStyleKey.itemKeys.add(
                        itemKeyClass.getConstructor(Map.class).newInstance(properties)
                );
            } catch (Exception e) {
                throw new RuntimeException(e);
            }
        });
        return cellStyleKey;
    }

    private static Map<Class<? extends ItemKey>, Map<String, String>> disassembleItemKey(String key) {
        final char[] charArray = StringUtils.deleteWhitespace(key).toCharArray();
        final StringBuilder itemKeyNameBuilder = new StringBuilder();
        final StringBuilder propertiesBuilder = new StringBuilder();
        final Map<Class<? extends ItemKey>, Map<String, String>> map = new HashMap<>();

        boolean inner = false;
        Class<? extends ItemKey> itemKeyClass = null;
        for (char c : charArray) {
            if (c == '{') {
                if(inner) {
                    throw new IllegalArgumentException("item key is valid");
                }
                inner = true;
                itemKeyClass = determineItemKey(itemKeyNameBuilder.toString());
                itemKeyNameBuilder.setLength(0);
            } else if (c == '}') {
                if(!inner) {
                    throw new IllegalArgumentException("item key is valid");
                }
                inner = false;
                map.put(itemKeyClass, disassembleProperties(propertiesBuilder.toString()));
                itemKeyClass = null;
                propertiesBuilder.setLength(0);
            } else if (inner) {
                propertiesBuilder.append(c);
            } else {
                itemKeyNameBuilder.append(c);
            }
        }
        return map;
    }

    private static Map<String, String> disassembleProperties(String properties) {
        final Map<String, String> map = new HashMap<>();
        final String[] parts = properties.split(";");
        for (String part : parts) {
            if (part.isEmpty()) {
                continue;
            }
            final String[] kv = part.split(":");
            if (kv.length == 1) {
                map.put(kv[0], part);
            } else {
                map.put(kv[0], kv[1]);
            }
        }
        return map;
    }

    private static Class<? extends ItemKey> determineItemKey(String itemKeyName) {
        for (Class<? extends ItemKey> itemKeyClass : ITEM_KEY_CLASSES) {
            final String[] prefixArray = itemKeyClass.getAnnotation(ItemKeySign.class).value();
            for (String prefix : prefixArray) {
                if (itemKeyName.equals(prefix)) {
                    return itemKeyClass;
                }
            }
        }
        throw new IllegalArgumentException("unknown item key \"" + itemKeyName + "\"");
    }

    public abstract static class ItemKey {

        public ItemKey(Map<String, String> properties) {

        }

        protected abstract void configureCellStyle(Workbook workbook, CellStyle cellStyle);

        protected final java.awt.Color getColorFromRGB(String rgbStr) {
            final int rgb = Integer.valueOf(rgbStr.substring(1), 16);
            int r = rgb >> 16;
            int g = (rgb & 0x00ff00) >> 8;
            int b = rgb & 0x0000ff;
            return new java.awt.Color(r, g, b);
        }
    }

    @Retention(RetentionPolicy.RUNTIME)
    @Target(ElementType.TYPE)
    public @interface ItemKeySign {

        String[] value();
    }

    // font {size: 16; bold; italic; family: 宋体; color: #ffaaee}
    @Getter
    @ItemKeySign({"font", "f"})
    protected static class FontKey extends ItemKey {

        private short heightInPoints;

        private XSSFColor xssfColor;

        private String fontName;

        private boolean italic;

        private boolean bold;

        public FontKey(Map<String, String> properties) {
            super(properties);
            if (properties.containsKey("size")) {
                heightInPoints = Short.parseShort(properties.get("size"));
            }
            if (properties.containsKey("color")) {
                xssfColor = new XSSFColor(getColorFromRGB(properties.get("color")));
            }
            if (properties.containsKey("family")) {
                fontName = properties.get("family");
            }
            if (properties.containsKey("italic")) {
                italic = Boolean.parseBoolean(properties.get("italic"));
            }
            if (properties.containsKey("bold")) {
                bold = Boolean.parseBoolean(properties.get("bold"));
            }
        }

        @Override
        public void configureCellStyle(Workbook workbook, CellStyle cellStyle) {
            final Font font = workbook.createFont();
            font.setFontName(fontName);
            font.setFontHeightInPoints(heightInPoints);
            font.setItalic(italic);
            font.setBold(bold);

            if (font instanceof XSSFFont) {
                ((XSSFFont) font).setColor(xssfColor);
            }
            cellStyle.setFont(font);
        }
    }

    // align {vertical: right; horizontal: center}
    @Getter
    @ItemKeySign({"align", "a"})
    protected static class AlignmentKey extends ItemKey {

        private HorizontalAlignment horizontalAlignment = HorizontalAlignment.CENTER;

        private VerticalAlignment verticalAlignment = VerticalAlignment.CENTER;

        public AlignmentKey(Map<String, String> properties) {
            super(properties);
            if (properties.containsKey("v")) {
                verticalAlignment = VerticalAlignment.valueOf(properties.get("v").toUpperCase());
            }
            if (properties.containsKey("vertical")) {
                verticalAlignment = VerticalAlignment.valueOf(properties.get("vertical").toUpperCase());
            }
            if (properties.containsKey("h")) {
                horizontalAlignment = HorizontalAlignment.valueOf(properties.get("h").toUpperCase());
            }
            if (properties.containsKey("horizontal")) {
                horizontalAlignment = HorizontalAlignment.valueOf(properties.get("horizontal").toUpperCase());
            }
        }

        @Override
        public void configureCellStyle(Workbook workbook, CellStyle cellStyle) {
            cellStyle.setAlignment(horizontalAlignment);
            cellStyle.setVerticalAlignment(verticalAlignment);
        }
    }

    // foreground {color: #aa1199; type: SOLID_FOREGROUND}
    @Getter
    @ItemKeySign({"foreground", "fg"})
    protected static class ForegroundKey extends ItemKey {

        private XSSFColor xssfColor;

        private FillPatternType fillPatternType = FillPatternType.SOLID_FOREGROUND;

        public ForegroundKey(Map<String, String> properties) {
            super(properties);
            if (properties.containsKey("color")) {
                xssfColor = new XSSFColor(getColorFromRGB(properties.get("color")));
            }
            if (properties.containsKey("type")) {
                fillPatternType = FillPatternType.valueOf(properties.get("type").toUpperCase());
            }
        }

        @Override
        public void configureCellStyle(Workbook workbook, CellStyle cellStyle) {
            if (cellStyle instanceof XSSFCellStyle) {
                XSSFCellStyle xssfCellStyle = (XSSFCellStyle) cellStyle;
                xssfCellStyle.setFillForegroundColor(xssfColor);
            }
            cellStyle.setFillPattern(fillPatternType);
        }
    }

    // border {style: thin; color: #bbaadd;}
    @Getter
    @ItemKeySign({"border", "b"})
    protected static class BorderKey extends ItemKey {

        private BorderStyle borderStyle = BorderStyle.THIN;

        private XSSFColor xssfColor;

        public BorderKey(Map<String, String> properties) {
            super(properties);
            if (properties.containsKey("color")) {
                xssfColor = new XSSFColor(getColorFromRGB(properties.get("color")));
            }
            if (properties.containsKey("style")) {
                borderStyle = BorderStyle.valueOf(properties.get("style").toUpperCase());
            }
        }

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
    }

    // dataFormat {format: 0.00%}
    @Getter
    @ItemKeySign({"dataFormat", "df"})
    protected static class DataFormatKey extends ItemKey {

        private String format = "@";

        public DataFormatKey(Map<String, String> properties) {
            super(properties);
            if (properties.containsKey("format")) {
                format = properties.get("format");
            }
        }

        @Override
        public void configureCellStyle(Workbook workbook, CellStyle cellStyle) {
            cellStyle.setDataFormat(workbook.createDataFormat().getFormat(format));
        }
    }
}
