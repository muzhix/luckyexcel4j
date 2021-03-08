package net.hanbd.luckyexcel4j.utils;

import cn.hutool.core.util.HexUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.XSSFColor;

import javax.annotation.Nullable;
import java.util.Objects;

/**
 * @author hanbd
 */
@Slf4j
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

    /**
     * 从{@link XSSFColor}中提取RGB hex颜色值. eg: {@code "#000000"}
     *
     * @param color
     * @return rgb hex or null
     */
    @Nullable
    public static String getRgbHexStr(@Nullable XSSFColor color) {
        String hexColor = null;
        if (Objects.nonNull(color)) {
            if (Objects.nonNull(color.getRGB())) {
                hexColor = "#" + HexUtil.encodeHexStr(color.getRGB());
            } else {
                // default black
                // TODO 边框线颜色如果为黑色,那么color.getRGB == null,是否要赋默认值？
//                hexColor = "#ffffff";
            }
        }
        return hexColor;
    }

    /**
     * eg: {@code "rgb(255, 0, 0)"}
     *
     * @param color
     * @return rgb str or null
     */
    @Nullable
    public static String getRgbStr(@Nullable XSSFColor color) {
        if (Objects.isNull(color)) {
            return null;
        }
        byte[] rgbBytes = color.getRGB();
        if (Objects.isNull(rgbBytes) || rgbBytes.length != 3) {
            return null;
        }

        StringBuilder sb = new StringBuilder();
        sb.append("rgb(");
        for (byte rgbByte : rgbBytes) {
            sb.append(rgbByte & 0xFF).append(",");
        }
        sb.deleteCharAt(sb.length() - 1);
        sb.append(")");
        return sb.toString();
    }
}
