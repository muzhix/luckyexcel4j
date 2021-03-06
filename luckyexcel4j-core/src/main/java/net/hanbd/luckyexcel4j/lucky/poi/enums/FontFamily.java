package net.hanbd.luckyexcel4j.lucky.poi.enums;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonEnumDefaultValue;
import com.fasterxml.jackson.annotation.JsonValue;
import lombok.AllArgsConstructor;
import lombok.Getter;

import java.util.Arrays;
import java.util.Map;
import java.util.Objects;
import java.util.function.Function;
import java.util.stream.Collectors;

/**
 * Luckysheet 支持的字体类型 TODO 扩展字体支持
 *
 * @author hanbd
 * @see
 * <a href="https://mengshukeji.gitee.io/LuckysheetDocs/zh/guide/cell.html#%E5%9F%BA%E6%9C%AC%E5%8D%95%E5%85%83%E6%A0%BC">
 * 基本单元格 ff</a>
 */
@AllArgsConstructor
public enum FontFamily {
    /**
     * Times New Roman
     */
    TIMES_NEW_ROMAN(0, "Times New Roman"),
    /**
     * Arial
     */
    @JsonEnumDefaultValue
    ARIAL(1, "Arial"),
    /**
     * Tahoma
     */
    TAHOMA(2, "Tahoma"),
    /**
     * Verdana
     */
    VERDANA(3, "Verdana"),
    /**
     * 微软雅黑
     */
    MS_YA_HEI(4, "微软雅黑"),
    /**
     * 宋体
     */
    SONG(5, "宋体"),
    /**
     * 黑体
     */
    ST_HEI_TI(6, "黑体"),
    /**
     * 楷体
     */
    ST_KAI_TI(7, "楷体"),
    /**
     * 仿宋
     */
    ST_FANG_SONG(8, "仿宋"),
    /**
     * 新宋体
     */
    ST_SONG(9, "新宋体"),
    /**
     * 华文新魏
     */
    HUA_WEN_XIN_WEI(10, "华文新魏"),
    /**
     * 华文行楷
     */
    HUA_WEN_XING_KAI(11, "华文行楷"),
    /**
     * 华文隶书
     */
    HUA_WEN_LI_SHU(12, "华文隶书");

    private final Integer type;
    @Getter
    private final String name;
    private static final Map<Integer, FontFamily> TYPES = Arrays.stream(FontFamily.values())
            .collect(Collectors.toMap(FontFamily::getType, Function.identity()));
    private static final Map<String, FontFamily> NAMES = Arrays.stream(FontFamily.values())
            .collect(Collectors.toMap(FontFamily::getName, Function.identity()));

    @JsonCreator
    public static FontFamily of(Integer type) {
        FontFamily fontFamily = TYPES.get(type);
        if (Objects.isNull(fontFamily)) {
            return FontFamily.ARIAL;
        }
        return fontFamily;
    }

    public static FontFamily of(String name) {
        FontFamily fontFamily = NAMES.get(name);
        if (Objects.isNull(fontFamily)) {
            return FontFamily.ARIAL;
        }
        return fontFamily;
    }

    @JsonValue
    public Integer getType() {
        return type;
    }
}
