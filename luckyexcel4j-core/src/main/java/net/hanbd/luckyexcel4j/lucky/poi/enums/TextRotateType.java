package net.hanbd.luckyexcel4j.lucky.poi.enums;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonEnumDefaultValue;
import com.fasterxml.jackson.annotation.JsonValue;
import lombok.AllArgsConstructor;

import java.util.Arrays;
import java.util.Map;
import java.util.function.Function;
import java.util.stream.Collectors;

/**
 * 字体旋转类型,<b>参考微软excel中 开始->对齐方式->方向</b>
 *
 * @author hanbd
 */
@AllArgsConstructor
public enum TextRotateType {
    /**
     * 不旋转
     */
    @JsonEnumDefaultValue
    NONE(0),
    /**
     * 逆时针45度
     */
    ANTICLOCKWISE_45(2),
    /**
     * 逆时针90度。 MS-excel 向上旋转文字
     */
    ANTICLOCKWISE_90(5),
    /**
     * 顺时针45度
     */
    CLOCKWISE_45(1),
    /**
     * 顺时针90度。MS-excel 向下旋转文字
     */
    CLOCKWISE_90(4),
    /**
     * 纵向。MS-excel 竖排文字
     */
    VERTICAL(3);

    private final Integer type;
    private static final Map<Integer, TextRotateType> TYPES = Arrays.stream(TextRotateType.values())
            .collect(Collectors.toMap(TextRotateType::getType, Function.identity()));

    @JsonCreator
    public static TextRotateType of(Integer type) {
        return TYPES.get(type);
    }

    @JsonValue
    public Integer getType() {
        return type;
    }


}
