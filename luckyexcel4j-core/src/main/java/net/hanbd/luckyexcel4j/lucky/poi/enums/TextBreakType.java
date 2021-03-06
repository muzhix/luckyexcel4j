package net.hanbd.luckyexcel4j.lucky.poi.enums;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonValue;
import lombok.AllArgsConstructor;

import java.util.Arrays;
import java.util.Map;
import java.util.function.Function;
import java.util.stream.Collectors;

/**
 * 文本换行类型
 *
 * @author hanbd
 */
@AllArgsConstructor
public enum TextBreakType {
    /**
     * 文本截断
     */
    TRUNCATION(0),
    /**
     * 文本溢出
     */
    OVERFLOW(1),
    /**
     * 自动换行
     */
    LINE_WRAP(2);

    private final Integer type;
    private static final Map<Integer, TextBreakType> TYPES = Arrays.stream(TextBreakType.values())
            .collect(Collectors.toMap(TextBreakType::getType, Function.identity()));

    @JsonCreator
    public static TextBreakType of(Integer type) {
        return TYPES.get(type);
    }

    @JsonValue
    public Integer getType() {
        return type;
    }
}
