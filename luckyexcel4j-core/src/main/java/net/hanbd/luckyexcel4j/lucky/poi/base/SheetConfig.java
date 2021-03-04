package net.hanbd.luckyexcel4j.lucky.poi.base;

import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;
import java.util.Map;

/**
 * LuckySheet config属性
 *
 * @author hanbd
 * @see <a
 * href="https://mengshukeji.gitee.io/LuckysheetDocs/zh/guide/sheet.html#config-rowlen">rowHeight</a>
 * @see <a
 * href="https://mengshukeji.gitee.io/LuckysheetDocs/zh/guide/sheet.html#config-columnlen">colWidth</a>
 * @see <a
 * href="https://mengshukeji.gitee.io/LuckysheetDocs/zh/guide/sheet.html#config-rowhidden">rowHidden</a>
 * @see <a
 * href="https://mengshukeji.gitee.io/LuckysheetDocs/zh/guide/sheet.html#config-colhidden">colHidden</a>
 */
@Data
@Builder
@AllArgsConstructor
@NoArgsConstructor
public class SheetConfig {
    /**
     * 合并单元格信息
     */
    private Map<String, MergeCell> merge;
    /**
     * 边框信息
     */
    private List<Border> borderInfo;
    /**
     * 行高信息. key: row index, value: row height
     */
    @JsonProperty("rowlen")
    private Map<Integer, Integer> rowHeight;
    /**
     * 列宽信息. key: column index, value: column width
     */
    @JsonProperty("columnlen")
    private Map<Integer, Integer> columnWidth;
    /**
     * 隐藏行. key: row index
     *
     * <p>key对应行被隐藏,value始终为0. TODO 使用HashSet优化,使其更符合java使用习惯
     */
    @JsonProperty("rowhidden")
    private Map<Integer, Integer> rowHidden;
    /**
     * 隐藏列. key: column index
     *
     * <p>key对应列被隐藏,value始终为0. TODO 使用HashSet优化,使其更符合java使用习惯
     */
    @JsonProperty("colhidden")
    private Map<Integer, Integer> columnHidden;


}
