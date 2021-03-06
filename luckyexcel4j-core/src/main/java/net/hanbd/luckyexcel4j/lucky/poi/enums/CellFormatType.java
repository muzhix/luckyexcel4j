package net.hanbd.luckyexcel4j.lucky.poi.enums;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonValue;
import lombok.AllArgsConstructor;
import org.apache.poi.ss.usermodel.CellType;

import java.util.Arrays;
import java.util.Map;
import java.util.function.Function;
import java.util.stream.Collectors;

/**
 * 单元格值Type类型
 *
 * @author hanbd
 * @see <a
 * href="https://mengshukeji.gitee.io/LuckysheetDocs/zh/guide/cell
 * .html#%E5%8F%AF%E9%80%89%E6%8B%A9%E7%9A%84%E8%AE%BE%E7%BD%AE%E5%A6%82%E4%B8%8B">
 * cellFormatType
 * </a>
 */
@AllArgsConstructor
public enum CellFormatType {
    /**
     * 自动格式
     */
    GENERAL("g"),
    /**
     * 纯文本字符串
     */
    STRING("s"),
    /**
     * 数字或货币
     */
    NUMBER("n"),
    /**
     * 日期时间
     */
    DATETIME("d");

    private final String type;
    private static final Map<String, CellFormatType> TYPES = Arrays.stream(CellFormatType.values())
            .collect(Collectors.toMap(CellFormatType::getType, Function.identity()));

    public static CellFormatType of(CellType cellType) {
        switch (cellType) {
            case NUMERIC:
                return NUMBER;
            case STRING:
            case FORMULA:
                return STRING;
            default:
                return GENERAL;

        }
    }

    @JsonCreator
    public static CellFormatType of(String type) {
        return TYPES.get(type);
    }

    @JsonValue
    public String getType() {
        return type;
    }
}
