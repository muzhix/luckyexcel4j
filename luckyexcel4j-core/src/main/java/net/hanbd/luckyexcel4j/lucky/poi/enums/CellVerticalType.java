package net.hanbd.luckyexcel4j.lucky.poi.enums;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonValue;
import lombok.AllArgsConstructor;

import java.util.Arrays;
import java.util.Map;
import java.util.function.Function;
import java.util.stream.Collectors;

/**
 * 单元格内值的垂直对齐方式
 *
 * @author hanbd
 */
@AllArgsConstructor
public enum CellVerticalType {
    /**
     * 居中对齐
     */
    CENTER(0),
    /**
     * 上对齐
     */
    TOP(1),
    /**
     * 下对齐
     */
    BOTTOM(2);

    private final Integer type;
    private static final Map<Integer, CellVerticalType> TYPES = Arrays.stream(CellVerticalType.values())
            .collect(Collectors.toMap(CellVerticalType::getType, Function.identity()));

    @JsonCreator
    public static CellVerticalType of(Integer type) {
        return TYPES.get(type);
    }

    @JsonValue
    public Integer getType() {
        return type;
    }
}
