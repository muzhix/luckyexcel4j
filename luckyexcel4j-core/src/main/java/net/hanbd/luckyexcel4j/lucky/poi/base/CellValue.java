package net.hanbd.luckyexcel4j.lucky.poi.base;

import com.fasterxml.jackson.annotation.JsonFormat;
import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
import net.hanbd.luckyexcel4j.lucky.poi.enums.CellHorizontalType;
import net.hanbd.luckyexcel4j.lucky.poi.enums.CellVerticalType;
import net.hanbd.luckyexcel4j.lucky.poi.enums.FontFamily;
import net.hanbd.luckyexcel4j.lucky.poi.enums.TextBreakType;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

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
@AllArgsConstructor
@Builder
@JsonInclude(JsonInclude.Include.NON_NULL)
public class CellValue {
    /**
     * 单元格值格式
     */
    @JsonProperty("ct")
    private CellFormat cellType;
    /**
     * 背景颜色Hex. eg: {@code #fff000}
     */
    @JsonProperty("bg")
    private String background;
    /**
     * 字体
     */
    @JsonProperty("ff")
    private FontFamily fontFamily;
    /**
     * 字体颜色. #fff000
     */
    @JsonProperty("fc")
    private String fontColor;
    /**
     * 粗体.
     */
    @JsonProperty("bl")
    @JsonFormat(shape = JsonFormat.Shape.NUMBER)
    private Boolean bold;
    /**
     * 斜体.
     */
    @JsonProperty("it")
    @JsonFormat(shape = JsonFormat.Shape.NUMBER)
    private Boolean italic;
    /**
     * 字体大小
     */
    @JsonProperty("fs")
    private Short fontsize;
    /**
     * 删除线.
     */
    @JsonProperty("cl")
    @JsonFormat(shape = JsonFormat.Shape.NUMBER)
    private Boolean cancelLine;
    /**
     * 垂直对齐. 0,居中; 1,上; 2,下
     */
    @JsonProperty("vt")
    private CellVerticalType verticalType;
    /**
     * 水平对齐. 0 居中、1 左、2右
     */
    @JsonProperty("ht")
    private CellHorizontalType horizontalType;
    /**
     * 合并单元格
     */
    @JsonProperty("mc")
    private MergeCell mergeCell;
    /**
     * 竖排文字,<b>参考微软excel中 开始->对齐方式->方向</b>
     * <p>
     * TODO 使用枚举
     *
     * <p>Text rotation,0: 0、1: 45 、2: -45、3 Vertical text、4: 90 、5: -90
     */
    @JsonProperty("tr")
    private Integer textRotate;
    /**
     * 文字旋转角度,取值范围[0-180].
     * <p>
     * 该字段与{@link CellValue#textRotate}其实表达的是一个意思,
     * {@link CellValue#textRotate}是{@link CellValue#rotateText}表意的子集。
     * 具体可查看{@link XSSFCellStyle#getRotation()}注释。
     *
     * @see XSSFCellStyle#getRotation()
     */
    @JsonProperty("rt")
    private Short rotateText;
    /**
     * 文本换行. 0 截断、1溢出、2 自动换行
     */
    @JsonProperty("tb")
    private TextBreakType textBreak;
    /**
     * 原始值.可能为字符串类型,也可能为数字类型. TODO 值优化处理
     */
    @JsonProperty("v")
    private String value;
    /**
     * 显示值
     */
    @JsonProperty("m")
    private String monitor;
    /**
     * 公式. eg: {@code "=SUM(233)" }
     */
    @JsonProperty("f")
    private String formula;
    /**
     * 批注
     */
    @JsonProperty("ps")
    private Comment comment;
}
