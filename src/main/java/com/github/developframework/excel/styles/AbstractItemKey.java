package com.github.developframework.excel.styles;

import java.util.Map;

/**
 * @author qiushui on 2023-05-22.
 */
public abstract class AbstractItemKey implements ItemKey {

    public AbstractItemKey(Map<String, String> properties) {

    }

    protected final java.awt.Color getColorFromRGB(String rgbStr) {
        final int rgb = Integer.valueOf(rgbStr.substring(1), 16);
        int r = rgb >> 16;
        int g = (rgb & 0x00ff00) >> 8;
        int b = rgb & 0x0000ff;
        return new java.awt.Color(r, g, b);
    }
}
