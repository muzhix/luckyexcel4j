package net.hanbd.luckyexcel4j.lucky.poi.base;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * 选中区域
 *
 * @author hanbd
 */
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class Selection {

    public Selection(Integer sheetIndex, Integer rowIndex, Integer columnIndex) {
        this.sheetIndex = sheetIndex;
        this.row = new StartEndPair(rowIndex, rowIndex);
        this.column = new StartEndPair(columnIndex, columnIndex);
    }

    /**
     * 工作表index
     */
    private Integer sheetIndex;
    /**
     * 行起止index范围
     */
    private StartEndPair row;
    /**
     * 列起止index范围
     */
    private StartEndPair column;

    /**
     * 起止index对
     */
    @AllArgsConstructor
    static class StartEndPair {
        /**
         * start index
         */
        private Integer start;
        /**
         * end index
         */
        private Integer end;
    }
}
