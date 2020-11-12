package net.hanbd.luckyexcel4j.lucky.poi.base;

import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

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
public class LuckySheetMerge {
    /**
     * 合并单元格主单元格行号
     */
    @JsonProperty("r")
    private Integer startRow;
    /**
     * 合并单元格主单元格列号
     */
    @JsonProperty("c")
    private Integer startCol;
    /**
     * 合并单元格所占行数
     */
    @JsonProperty("rs")
    private Integer rowsNum;
    /**
     * 合并单元格所占列数
     */
    @JsonProperty("cs")
    private Integer colsNum;

    public String mainCellStr() {
        return startRow + "_" + startCol;
    }
}
