package net.hanbd.luckyexcel4j.lucky.poi.base;

import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * 公式链
 *
 * @author hanbd
 * @see <a
 * href="https://mengshukeji.gitee.io/LuckysheetDocs/zh/guide/sheet.html#calcchain">calcChain</a>
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class SheetCalcChain {

    @JsonProperty("r")
    private Integer row;
    @JsonProperty("c")
    private Integer column;
    /**
     * {@link SheetMeta#getIndex()}
     */
    @JsonProperty("index")
    private Integer sheetIndex;
    private Integer
}
