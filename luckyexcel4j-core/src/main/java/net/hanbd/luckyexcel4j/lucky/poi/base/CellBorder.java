package net.hanbd.luckyexcel4j.lucky.poi.base;

import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.*;
import net.hanbd.luckyexcel4j.lucky.poi.enums.BorderRangeType;
import net.hanbd.luckyexcel4j.lucky.poi.enums.BorderStyle;

/**
 * 单元格边框. {@link BorderRangeType#CELL}
 *
 * @author hanbd
 * @see <a href="https://mengshukeji.gitee.io/LuckysheetDocs/zh/guide/sheet.html#config-borderinfo">
 * borderInfo</a>
 */
@EqualsAndHashCode(callSuper = true)
@Data
public class CellBorder extends Border {

    private Value value;

    /**
     * 新建对象,并初始化相应值
     */
    public CellBorder() {
        this.rangeType = BorderRangeType.CELL;
        this.value = new Value();
    }

    @Data
    @NoArgsConstructor
    static class Value {
        /**
         * 单元格行索引值
         */
        @JsonProperty("row_index")
        private Integer row;
        /**
         * 单元格列索引值
         */
        @JsonProperty("col_index")
        private Integer col;
        /**
         * 左侧边框样式
         */
        @JsonProperty("l")
        private Style left;
        /**
         * 右侧边框样式
         */
        @JsonProperty("r")
        private Style right;
        /**
         * 顶部边框样式
         */
        @JsonProperty("t")
        private Style top;
        /**
         * 底部边框样式
         */
        @JsonProperty("b")
        private Style bottom;
    }

    @Data
    @AllArgsConstructor
    @Builder
    static class Style {
        /**
         * 边框线格式 {@link BorderStyle#getStyle}
         */
        private Integer style;
        /**
         * 颜色. 格式: {@code rgb(255,0,0)}
         */
        private String color;
    }
}
