package net.hanbd.luckyexcel4j.lucky.poi.enums;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonValue;
import lombok.AllArgsConstructor;

import java.util.Arrays;
import java.util.Map;
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
    TIMES_NEW_ROMAN(0),
    /**
     * Arial
     */
    ARIAL(1),
    /**
     * Tahoma
     */
    TAHOMA(2),
    /**
     * Verdana
     */
    VERDANA(3),
    /**
     * 微软雅黑
     */
    MS_YA_HEI(4),
    /**
     * 宋体
     */
    SONG(5),
    /**
     * 黑体
     */
    ST_HEI_TI(6),
    /**
     * 楷体
     */
    ST_KAI_TI(7),
    /**
     * 仿宋
     */
    ST_FANG_SONG(8),
    /**
     * 新宋体
     */
    ST_SONG(9),
    /**
     * 华文新魏
     */
    HUA_WEN_XIN_WEI(10),
    /**
     * 华文行楷
     */
    HUA_WEN_XING_KAI(11),
    /**
     * 华文隶书
     */
    HUA_WEN_LI_SHU(12);

    private final Integer type;
    private static final Map<Integer, FontFamily> TYPES = Arrays.stream(FontFamily.values())
            .collect(Collectors.toMap(FontFamily::getType, Function.identity()));

    @JsonCreator
    public static FontFamily of(Integer type) {
        return TYPES.get(type);
    }

    @JsonValue
    public Integer getType() {
        return type;
    }
}
