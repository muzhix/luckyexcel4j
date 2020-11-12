package net.hanbd.luckyexcel4j.lucky.poi.base;

import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
import net.hanbd.luckyexcel4j.lucky.poi.enums.LuckySheetCellFormatType;

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
public class LuckySheetCellFormat {
    /**
     * Format格式的定义字符串
     */
    @JsonProperty("fa")
    private String format;
    /**
     * Type类型 {@link LuckySheetCellFormatType#getType}
     */
    @JsonProperty("t")
    private LuckySheetCellFormatType type;
}
