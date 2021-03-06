package net.hanbd.luckyexcel4j.lucky.poi.base;

import com.fasterxml.jackson.annotation.JsonIgnore;
import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import javax.validation.constraints.Min;
import java.util.Objects;

/**
 * luckysheet中合并单元格的表示
 * <p>
 * +-----------------+       +----------------+
 * |(0,0)   |(0,1)   |       |                |
 * +-----------------+ +---> |                |
 * |(1,0)   |(1,1)   |       |                |
 * +-----------------+       +----------------+
 * <p>
 * 如上4个单元格合并后: startRow=0,startCol=0,rowsNum=2,colsNum=2
 *
 * @author hanbd
 * @see <a
 * href="https://mengshukeji.gitee.io/LuckysheetDocs/zh/guide/cell.html#%E5%90%88%E5%B9%B6%E5%8D%95%E5%85%83%E6%A0%BC">
 * luckysheet合并单元格 </a>
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class MergeCell implements IMergeCell {
    /**
     * 合并单元格主单元格行号, <b>0 based</b>
     */
    @JsonProperty("r")
    @Min(0)
    private Integer startRow;
    /**
     * 合并单元格主单元格列号, <b>0 based</b>
     */
    @JsonProperty("c")
    @Min(0)
    private Integer startCol;
    /**
     * 合并单元格所占行数
     */
    @JsonProperty("rs")
    @Min(1)
    private Integer rowsNum;
    /**
     * 合并单元格所占列数
     */
    @JsonProperty("cs")
    @Min(1)
    private Integer colsNum;

    @JsonIgnore
    private MainCell mainCell;

    /**
     * 获得合并单元格的定位字符串(左上角单元格).
     *
     * @return
     */
    public String mainCellStr() {
        return startRow + "_" + startCol;
    }

    /**
     * 获得合并单元格的主单元格
     *
     * @return
     */
    public MainCell mainCell() {
        if (Objects.isNull(mainCell)) {
            mainCell = new MainCell(startRow, startCol);
        }
        return mainCell;
    }

    /**
     * 合并单元格的主单元格
     */
    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    @Deprecated
    static class MainCell implements IMergeCell {
        /**
         * 合并单元格主单元格行号, <b>0 based</b>
         */
        @JsonProperty("r")
        protected Integer row;
        /**
         * 合并单元格主单元格列号, <b>0 based</b>
         */
        @JsonProperty("c")
        protected Integer column;
    }
}
