package net.hanbd.luckyexcel4j.lucky.poi.base;

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * 单元格数据
 *
 * @author hanbd
 * @see <a
 * href="https://mengshukeji.gitee.io/LuckysheetDocs/zh/guide/cell.html#%E5%9F%BA%E6%9C%AC%E5%8D%95%E5%85%83%E6%A0%BC">
 * 单元格</a>
 */
@Data
@NoArgsConstructor
@JsonInclude(JsonInclude.Include.NON_NULL)
public class LuckySheetCellValue {
    /**
     * 单元格值格式
     */
    @JsonProperty("ct")
    private LuckySheetCellFormat cellType;
    /**
     * 背景颜色. eg: {@code #fff000}
     */
    @JsonProperty("bg")
    private String background;
    /**
     * 字体
     */
    @JsonProperty("ff")
    private String fontFamily;
    /**
     * 字体颜色. #fff000
     */
    @JsonProperty("fc")
    private String fontColor;
    /**
     * 字体大小
     */
    @JsonProperty("fs")
    private Integer fontsize;
    /**
     * 粗体. 0,常规; 1,加粗
     */
    @JsonProperty("bl")
    private Integer bold;
    /**
     * 斜体. 0,常规; 1,斜体
     */
    @JsonProperty("it")
    private Integer italic;
    /**
     * 删除线. 0,常规; 1,删除线
     */
    @JsonProperty("cl")
    private Integer cancelLine;
    /**
     * 垂直对齐. 0,居中; 1,上; 2,下
     */
    @JsonProperty("vt")
    private Integer verticalType;
    /**
     * 水平对齐. 0 居中、1 左、2右
     */
    @JsonProperty("ht")
    private Integer horizontalType;
    /**
     * 合并单元格
     */
    @JsonProperty("mc")
    private LuckySheetMerge mergeCell;
    /**
     * 竖排文字
     *
     * <p>Text rotation,0: 0、1: 45 、2: -45、3 Vertical text、4: 90 、5: -90
     */
    @JsonProperty("tr")
    private Integer textRotate;
    /**
     * 文字旋转角度,取值范围[0-180]
     */
    @JsonProperty("rt")
    private Integer rotateText;
    /**
     * 文本换行. 0 截断、1溢出、2 自动换行 TODO 优化为enum
     */
    @JsonProperty("tb")
    private Integer textBreak;
    /**
     * 原始值
     */
    @JsonProperty("v")
    private String value;
    /**
     * 显示值
     */
    @JsonProperty("m")
    private String monitor;
    /**
     * 公式
     */
    @JsonProperty("f")
    private String formula;

    // TODO 支持批注
}
