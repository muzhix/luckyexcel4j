package net.hanbd.luckyexcel4j.lucky.poi.enums;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonValue;
import lombok.AllArgsConstructor;
import lombok.Getter;

import java.util.Arrays;
import java.util.Map;
import java.util.function.Function;
import java.util.stream.Collectors;

/**
 * 边框线格式
 *
 * @author hanbd
 * @see <a
 * href="https://mengshukeji.gitee.io/LuckysheetDocs/zh/guide/sheet.html#config-borderinfo">config-borderInfo</a>
 */
@AllArgsConstructor
public enum BorderStyleType {
    /**
     * 无
     */
    NONE(0, "None", org.apache.poi.ss.usermodel.BorderStyle.NONE),
    /**
     * Thin
     */
    THIN(1, "Thin", org.apache.poi.ss.usermodel.BorderStyle.THIN),
    /**
     * Hair
     */
    HAIR(2, "Hair", org.apache.poi.ss.usermodel.BorderStyle.HAIR),
    /**
     * Dotted
     */
    DOTTED(3, "Dotted", org.apache.poi.ss.usermodel.BorderStyle.DOTTED),
    /**
     * Dashed
     */
    DASHED(4, "Dashed", org.apache.poi.ss.usermodel.BorderStyle.DASHED),
    /**
     * DashDot
     */
    DASH_DOT(5, "DashDot", org.apache.poi.ss.usermodel.BorderStyle.DASH_DOT),
    /**
     * DashDotDot
     */
    DASH_DOT_DOT(6, "DashDotDot", org.apache.poi.ss.usermodel.BorderStyle.DASH_DOT_DOT),
    /**
     * Double
     */
    DOUBLE(7, "Double", org.apache.poi.ss.usermodel.BorderStyle.DOUBLE),
    /**
     * Medium
     */
    MEDIUM(8, "Medium", org.apache.poi.ss.usermodel.BorderStyle.MEDIUM),
    /**
     * MediumDashed
     */
    MEDIUM_DASHED(9, "MediumDashed", org.apache.poi.ss.usermodel.BorderStyle.MEDIUM_DASHED),
    /**
     * MediumDashDot
     */
    MEDIUM_DASH_DOT(10, "MediumDashDot", org.apache.poi.ss.usermodel.BorderStyle.MEDIUM_DASH_DOT),
    /**
     * MediumDashDotDot
     */
    MEDIUM_DASH_DOT_DOT(11, "MediumDashDotDot", org.apache.poi.ss.usermodel.BorderStyle.MEDIUM_DASH_DOT_DOT),
    /**
     * SlantedDashDot
     */
    SLANTED_DASH_DOT(12, "SlantedDashDot", org.apache.poi.ss.usermodel.BorderStyle.SLANTED_DASH_DOT),
    /**
     * Thick
     */
    THICK(13, "Thick", org.apache.poi.ss.usermodel.BorderStyle.THICK);

    /**
     * 格式index
     */
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
    private final org.apache.poi.ss.usermodel.BorderStyle poiStyle;
    private static final Map<Integer, BorderStyleType> STYLES = Arrays.stream(BorderStyleType.values())
            .collect(Collectors.toMap(BorderStyleType::getStyle, Function.identity()));

    public static BorderStyleType valueOf(org.apache.poi.ss.usermodel.BorderStyle poiStyle) {
        return valueOf(poiStyle.name());
    }

    @JsonCreator
    public static BorderStyleType of(String intStyle) {
        return STYLES.get(Integer.valueOf(intStyle));
    }

    public static BorderStyleType of(org.apache.poi.ss.usermodel.BorderStyle poiStyle) {
        return valueOf(poiStyle.name());
    }

    @JsonValue
    public Integer getStyle() {
        return style;
    }
}
