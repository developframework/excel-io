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
import java.util.*;
import java.util.stream.Collectors;

/**
 * @author qiushui on 2023-05-19.
 */
public class CellStyleKey implements ItemKey {

    @SuppressWarnings("unchecked")
    private static final Class<? extends ItemKey>[] ITEM_KEY_CLASSES = new Class[]{
            AlignmentKey.class,
            BorderKey.class,
            ConfigKey.class,
            DataFormatKey.class,
            FontKey.class,
            ForegroundKey.class
    };

    private final List<ItemKey> itemKeys = new LinkedList<>();

    @Override
    public void configureCellStyle(Workbook workbook, CellStyle cellStyle) {
        for (ItemKey itemKey : itemKeys) {
            itemKey.configureCellStyle(workbook, cellStyle);
        }
    }

    @Override
    public String toString() {
        return itemKeys
                .stream()
                .sorted(Comparator.comparing(ik -> ik.getClass().getSimpleName()))
                .map(ItemKey::toString)
                .filter(s -> !s.endsWith("{}"))
                .collect(Collectors.joining(" "));
    }

    public static boolean isCellStyleKey(String key) {
        return key.isEmpty() || key.matches("^(\\s*\\w+\\s*\\{(.+?)*}\\s*)+$");
    }

    public static CellStyleKey parse(String key) {
        final CellStyleKey cellStyleKey = new CellStyleKey();
        final Map<Class<? extends ItemKey>, Map<String, String>> itemKeyMetadataMap = disassembleItemKey(key);

        for (Class<? extends ItemKey> itemKeyClass : ITEM_KEY_CLASSES) {
            final Map<String, String> properties = itemKeyMetadataMap.getOrDefault(itemKeyClass, Collections.emptyMap());
            try {
                cellStyleKey.itemKeys.add(
                        itemKeyClass.getConstructor(Map.class).newInstance(properties)
                );
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        return cellStyleKey;
    }

    private static Map<Class<? extends ItemKey>, Map<String, String>> disassembleItemKey(String key) {
        final char[] charArray = key.toCharArray();
        final StringBuilder itemKeyNameBuilder = new StringBuilder();
        final StringBuilder propertiesBuilder = new StringBuilder();
        final Map<Class<? extends ItemKey>, Map<String, String>> map = new HashMap<>();

        boolean inner = false;
        boolean quote = false;
        Class<? extends ItemKey> itemKeyClass = null;
        for (char c : charArray) {
            switch (c) {
                case ' ':
                case '\t':
                case '\n': {
                    if (quote) {
                        propertiesBuilder.append(c);
                    } else {
                        continue;
                    }
                }
                break;
                case '\'': {
                    quote = !quote;
                    propertiesBuilder.append(c);
                }
                break;
                case '{': {
                    if (inner) {
                        throw new IllegalArgumentException("item key is valid");
                    }
                    inner = true;
                    itemKeyClass = determineItemKey(itemKeyNameBuilder.toString());
                    itemKeyNameBuilder.setLength(0);
                }
                break;
                case '}': {
                    if (!inner) {
                        throw new IllegalArgumentException("item key is valid");
                    }
                    inner = false;
                    map.put(itemKeyClass, disassembleProperties(propertiesBuilder.toString()));
                    itemKeyClass = null;
                    propertiesBuilder.setLength(0);
                }
                break;
                default: {
                    if (inner) {
                        propertiesBuilder.append(c);
                    } else {
                        itemKeyNameBuilder.append(c);
                    }
                }
                break;
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
            final int i = part.indexOf(":");
            if (i == -1) {
                map.put(part, "true");
            } else {
                map.put(part.substring(0, i), StringUtils.strip(part.substring(i + 1), "'"));
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

    @Retention(RetentionPolicy.RUNTIME)
    @Target(ElementType.TYPE)
    public @interface ItemKeySign {

        String[] value();
    }

    // font {size: 16; bold; italic; family: 宋体; color: #ffaaee}
    @Getter
    @ItemKeySign({"font", "f"})
    protected static class FontKey extends AbstractItemKey {

        private short size;

        private XSSFColor xssfColor;

        private String color;

        private String family;

        private boolean italic;

        private boolean bold;

        public FontKey(Map<String, String> properties) {
            super(properties);
            if (properties.containsKey("size")) {
                size = Short.parseShort(properties.get("size"));
            }
            if (properties.containsKey("color")) {
                color = properties.get("color");
                xssfColor = new XSSFColor(getColorFromRGB(color));
            }
            if (properties.containsKey("family")) {
                family = properties.get("family");
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
            if (family != null) {
                font.setFontName(family);
            }
            if (size != 0) {
                font.setFontHeightInPoints(size);
            }
            font.setItalic(italic);
            font.setBold(bold);

            if(xssfColor != null) {
                if (font instanceof XSSFFont) {
                    ((XSSFFont) font).setColor(xssfColor);
                }
            }
            cellStyle.setFont(font);
        }

        @Override
        public String toString() {
            List<String> list = new LinkedList<>();
            if (size != 0) {
                list.add("size: " + size);
            }
            if (color != null) {
                list.add("color: " + color);
            }
            if (family != null) {
                list.add("family: '" + family + "'");
            }
            if (italic) {
                list.add("italic");
            }
            if (bold) {
                list.add("bold");
            }
            return String.format("font {%s}", StringUtils.join(list, "; "));
        }
    }

    // align {vertical: right; horizontal: center}
    @Getter
    @ItemKeySign({"align", "a"})
    protected static class AlignmentKey extends AbstractItemKey {

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

        @Override
        public String toString() {
            List<String> list = new LinkedList<>();
            if (verticalAlignment != VerticalAlignment.CENTER) {
                list.add("vertical: " + verticalAlignment);
            }
            if (horizontalAlignment != HorizontalAlignment.CENTER) {
                list.add("horizontal: " + horizontalAlignment);
            }
            return String.format("align {%s}", StringUtils.join(list, "; "));
        }
    }

    // foreground {color: #aa1199; type: SOLID_FOREGROUND}
    @Getter
    @ItemKeySign({"foreground", "fg"})
    protected static class ForegroundKey extends AbstractItemKey {

        private XSSFColor xssfColor;

        private String color;

        private FillPatternType type = FillPatternType.SOLID_FOREGROUND;

        public ForegroundKey(Map<String, String> properties) {
            super(properties);
            if (properties.containsKey("color")) {
                color = properties.get("color");
                xssfColor = new XSSFColor(getColorFromRGB(color));
            }
            if (properties.containsKey("type")) {
                type = FillPatternType.valueOf(properties.get("type").toUpperCase());
            }
        }

        @Override
        public void configureCellStyle(Workbook workbook, CellStyle cellStyle) {
            if(xssfColor != null) {
                if (cellStyle instanceof XSSFCellStyle) {
                    XSSFCellStyle xssfCellStyle = (XSSFCellStyle) cellStyle;
                    xssfCellStyle.setFillForegroundColor(xssfColor);
                }
                cellStyle.setFillPattern(type);
            }
        }

        @Override
        public String toString() {
            List<String> list = new LinkedList<>();
            if (color != null) {
                list.add("color: " + color);
            }
            if (type != FillPatternType.SOLID_FOREGROUND) {
                list.add("type: " + type);
            }
            return String.format("foreground {%s}", StringUtils.join(list, "; "));
        }
    }

    // border {style: thin; color: #bbaadd;}
    @Getter
    @ItemKeySign({"border", "b"})
    protected static class BorderKey extends AbstractItemKey {

        private BorderStyle style = BorderStyle.THIN;

        private String color;

        private XSSFColor xssfColor;

        public BorderKey(Map<String, String> properties) {
            super(properties);
            if (properties.containsKey("color")) {
                color = properties.get("color");
                xssfColor = new XSSFColor(getColorFromRGB(color));
            }
            if (properties.containsKey("style")) {
                style = BorderStyle.valueOf(properties.get("style").toUpperCase());
            }
        }

        @Override
        public void configureCellStyle(Workbook workbook, CellStyle cellStyle) {
            cellStyle.setBorderTop(style);
            cellStyle.setBorderRight(style);
            cellStyle.setBorderBottom(style);
            cellStyle.setBorderLeft(style);

            if(xssfColor != null) {
                if (cellStyle instanceof XSSFCellStyle) {
                    XSSFCellStyle xssfCellStyle = (XSSFCellStyle) cellStyle;
                    xssfCellStyle.setTopBorderColor(xssfColor);
                    xssfCellStyle.setRightBorderColor(xssfColor);
                    xssfCellStyle.setBottomBorderColor(xssfColor);
                    xssfCellStyle.setLeftBorderColor(xssfColor);
                }
            }
        }

        @Override
        public String toString() {
            List<String> list = new LinkedList<>();
            if (style != BorderStyle.THIN) {
                list.add("style: " + style);
            }
            if (color != null) {
                list.add("color: " + color);
            }
            return String.format("border {%s}", StringUtils.join(list, "; "));
        }
    }

    // dataFormat {format: 0.00%}
    @Getter
    @ItemKeySign({"dataFormat", "df"})
    protected static class DataFormatKey extends AbstractItemKey {

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

        @Override
        public String toString() {
            List<String> list = new LinkedList<>();
            if (!format.equals("@")) {
                list.add("format: '" + format + "'");
            }
            return String.format("dataFormat {%s}", StringUtils.join(list, "; "));
        }
    }

    // config {wrapText}
    @Getter
    @ItemKeySign({"config", "c"})
    protected static class ConfigKey extends AbstractItemKey {

        private boolean wrapText;

        public ConfigKey(Map<String, String> properties) {
            super(properties);
            if (properties.containsKey("wrapText")) {
                wrapText = Boolean.parseBoolean(properties.get("wrapText"));
            }
        }

        @Override
        public void configureCellStyle(Workbook workbook, CellStyle cellStyle) {
            cellStyle.setWrapText(wrapText);
        }

        @Override
        public String toString() {
            List<String> list = new LinkedList<>();
            if (wrapText) {
                list.add("wrapText");
            }
            return String.format("config {%s}", StringUtils.join(list, "; "));
        }
    }
}
