package net.hanbd.luckyexcel4j.utils;

/**
 * @author hanbd
 */
public class PoiUtil {
    /**
     * ratio, default 0.75 1in = 2.54cm = 25.4mm = 72pt = 6pc, pt = 1/72 In, px = 1/dpi In
     */
    public static final float POINT_TO_PIXEL_RATIO_BY_DPI = 72F / 96;

    /**
     * 获得像素列宽 TODO 确认具体算法公式
     *
     * @param columnWidth
     * @return
     * @see {@link org.apache.poi.xssf.usermodel.XSSFSheet#getColumnWidthInPixels(int)}
     */
    public static Integer getColumnWidthPixel(Double columnWidth) {
        return (int) Math.round((columnWidth - 0.83) * 8 + 5);
    }

    public static Integer getRowHeightPixel(Double rowHeightPoint) {
        return (int) Math.round(rowHeightPoint / POINT_TO_PIXEL_RATIO_BY_DPI);
    }
}
