package net.hanbd.luckyexcel4j.lucky.poi.enums;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonValue;
import lombok.AllArgsConstructor;

import java.util.Arrays;
import java.util.Map;
import java.util.function.Function;
import java.util.stream.Collectors;

/**
 * 单元格的水平对齐方式
 *
 * @author hanbd
 */
@AllArgsConstructor
public enum CellHorizontalType {
    /**
     * 居中对齐
     */
    CENTER(0),
    /**
     * 左对齐
     */
    LEFT(1),
    /**
     * 右对齐
     */
    RIGHT(2);

    private final Integer type;
    private static final Map<Integer, CellHorizontalType> TYPES = Arrays.stream(CellHorizontalType.values())
            .collect(Collectors.toMap(CellHorizontalType::getType, Function.identity()));

    @JsonCreator
    public static CellHorizontalType of(Integer type) {
        return TYPES.get(type);
    }

    @JsonValue
    public Integer getType() {
        return type;
    }
}
