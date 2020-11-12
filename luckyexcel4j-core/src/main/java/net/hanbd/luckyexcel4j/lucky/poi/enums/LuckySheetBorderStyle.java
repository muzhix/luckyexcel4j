package net.hanbd.luckyexcel4j.lucky.poi.enums;

import lombok.AllArgsConstructor;
import lombok.Getter;
import net.hanbd.luckyexcel4j.lucky.poi.base.LuckySheetBorder;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.xwpf.usermodel.Borders;

/**
 * 边框线格式
 *
 * @author hanbd
 * @see <a
 * href="https://mengshukeji.gitee.io/LuckysheetDocs/zh/guide/sheet.html#config-borderinfo">config-borderInfo</a>
 */
@AllArgsConstructor
public enum LuckySheetBorderStyle {
    /**
     * 无
     */
    NONE(0, "None", BorderStyle.NONE),
    /**
     * Thin
     */
    THIN(1, "Thin", BorderStyle.THIN),
    /**
     * Hair
     */
    HAIR(2, "Hair", BorderStyle.HAIR),
    /**
     * Dotted
     */
    DOTTED(3, "Dotted", BorderStyle.DOTTED),
    /**
     * Dashed
     */
    DASHED(4, "Dashed", BorderStyle.DASHED),
    /**
     * DashDot
     */
    DASH_DOT(5, "DashDot", BorderStyle.DASH_DOT),
    /**
     * DashDotDot
     */
    DASH_DOT_DOT(6, "DashDotDot", BorderStyle.DASH_DOT_DOT),
    /**
     * Double
     */
    DOUBLE(7, "Double", BorderStyle.DOUBLE),
    /**
     * Medium
     */
    MEDIUM(8, "Medium", BorderStyle.MEDIUM),
    /**
     * MediumDashed
     */
    MEDIUM_DASHED(9, "MediumDashed", BorderStyle.MEDIUM_DASHED),
    /**
     * MediumDashDot
     */
    MEDIUM_DASH_DOT(10, "MediumDashDot", BorderStyle.MEDIUM_DASH_DOT),
    /**
     * MediumDashDotDot
     */
    MEDIUM_DASH_DOT_DOT(11, "MediumDashDotDot", BorderStyle.MEDIUM_DASH_DOT_DOT),
    /**
     * SlantedDashDot
     */
    SLANTED_DASH_DOT(12, "SlantedDashDot", BorderStyle.SLANTED_DASH_DOT),
    /**
     * Thick
     */
    THICK(13, "Thick", BorderStyle.THICK);

    /**
     * 格式index
     */
    @Getter
    private final Integer style;
    /**
     * 格式名
     */
    @Getter
    private final String name;
    /**
     * 对应的poi border style
     */
    @Getter
    private final BorderStyle poiStyle;

    public static LuckySheetBorderStyle valueOf(BorderStyle poiStyle) {
        return valueOf(poiStyle.name());
    }
}
