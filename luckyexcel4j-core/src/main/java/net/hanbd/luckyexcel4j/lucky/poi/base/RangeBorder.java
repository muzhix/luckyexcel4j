package net.hanbd.luckyexcel4j.lucky.poi.base;

import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.Builder;
import lombok.Data;
import lombok.EqualsAndHashCode;
import net.hanbd.luckyexcel4j.lucky.poi.enums.BorderRangeType;
import net.hanbd.luckyexcel4j.lucky.poi.enums.BorderStyle;
import net.hanbd.luckyexcel4j.lucky.poi.enums.BorderType;

import java.util.List;

/**
 * 区域边框. {@link BorderRangeType#RANGE}
 *
 * @author hanbd
 * @see <a href="https://mengshukeji.gitee.io/LuckysheetDocs/zh/guide/sheet.html#config-borderinfo">
 * borderInfo</a>
 */
@EqualsAndHashCode(callSuper = true)
@Data
@Builder
public class RangeBorder extends Border {

    /**
     * 边框线位置类型
     */
    private BorderType borderType;
    /**
     * 边框线类型
     */
    private BorderStyle style;
    /**
     * 边框线颜色HEX. eg: #0000ff
     */
    private String color;
    /**
     * 区域范围
     */
    @JsonProperty("range")
    private List<Range> ranges;
}
