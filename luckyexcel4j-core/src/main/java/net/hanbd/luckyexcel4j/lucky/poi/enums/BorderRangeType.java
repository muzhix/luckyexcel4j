package net.hanbd.luckyexcel4j.lucky.poi.enums;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonValue;
import lombok.AllArgsConstructor;

import java.util.Arrays;
import java.util.Map;
import java.util.function.Function;
import java.util.stream.Collectors;

/**
 * 边框范围类型
 *
 * @author hanbd
 * @see <a href="https://mengshukeji.gitee.io/LuckysheetDocs/zh/guide/sheet.html#config-borderinfo">border=info</a>
 */
@AllArgsConstructor
public enum BorderRangeType {
    /**
     * 区域
     */
    RANGE("range"),
    /**
     * 单个单元格
     */
    CELL("cell");
    private final String type;
    private static final Map<String, BorderRangeType> TYPES = Arrays.stream(BorderRangeType.values())
            .collect(Collectors.toMap(BorderRangeType::getType, Function.identity()));

    @JsonCreator
    public static BorderRangeType of(String rangeType) {
        return TYPES.get(rangeType);
    }

    @JsonValue
    public String getType() {
        return type;
    }
}
