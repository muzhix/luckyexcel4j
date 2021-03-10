package net.hanbd.luckyexcel4j.lucky.poi.base;

import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.EqualsAndHashCode;
import net.hanbd.luckyexcel4j.lucky.poi.enums.BorderRangeType;
import net.hanbd.luckyexcel4j.lucky.poi.enums.BorderStyleType;
import net.hanbd.luckyexcel4j.lucky.poi.enums.BorderPosType;

import java.util.Collections;
import java.util.List;

/**
 * 区域边框. {@link BorderRangeType#RANGE}
 * <p>
 * 其实在从excel -> luckyJson 的转换过程中该类没用, ECMA-376规范中边框都是绑定在单元格上,
 * 只要所有单元格的边框正常解析,最后就能正常显示。
 * <p>
 * 在从luckyJson -> excel的转换中,需要将luckySheet设置的区域边框转成单元格边框集合,
 * 然后设置到对应单元格上,该转换过程由方法{@link RangeBorder#toCellBorders()}来实现
 *
 * @author hanbd
 * @see <a href="https://mengshukeji.gitee.io/LuckysheetDocs/zh/guide/sheet.html#config-borderinfo">
 * borderInfo</a>
 */
@EqualsAndHashCode(callSuper = true)
@Data
@AllArgsConstructor
@Builder
public class RangeBorder extends Border {

    public RangeBorder() {
        rangeType = BorderRangeType.RANGE;
    }

    /**
     * 边框线位置类型
     */
    private BorderPosType borderType;
    /**
     * 边框线类型
     */
    private BorderStyleType style;
    /**
     * 边框线颜色HEX. eg: #0000ff
     */
    private String color;
    /**
     * 区域范围
     */
    @JsonProperty("range")
    private List<Range> ranges;

    /**
     * TODO
     *
     * @return
     */
    public List<CellBorder> toCellBorders() {
        return Collections.emptyList();
    }
}
