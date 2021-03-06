package net.hanbd.luckyexcel4j.lucky.poi.base;

import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
import net.hanbd.luckyexcel4j.lucky.poi.enums.CellFormatType;

/**
 * 单元格值格式
 *
 * @author hanbd
 * @see <a
 * href="https://mengshukeji.gitee.io/LuckysheetDocs/zh/guide/cell
 * .html#%E5%8D%95%E5%85%83%E6%A0%BC%E5%80%BC%E6%A0%BC%E5%BC%8F">
 * 单元格值格式</a>
 */
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class CellFormat {
    /**
     * Format格式的定义字符串
     * TODO 优化format的赋值问题
     */
    @JsonProperty("fa")
    private String format;
    /**
     * Type类型 {@link CellFormatType#getType}
     */
    @JsonProperty("t")
    private CellFormatType type;

    /**
     * format 类型,TODO 使用接口实现或类继承实现
     */
    static class Format {
        /**
         * 自动
         */
        public static final String GENERAL = "General";
        /**
         * 纯文本
         */
        public static final String TEXT = "@";
        /**
         * 整数
         */
        public static final String NUMBER_INT = "0";
        /**
         * 数字,一位小数
         */
        public static final String NUMBER_DOT_ONE = "0.0";
        /**
         * 数字,两位小数
         */
        public static final String NUMBER_DOT_TWO = "0.00";
    }
}
