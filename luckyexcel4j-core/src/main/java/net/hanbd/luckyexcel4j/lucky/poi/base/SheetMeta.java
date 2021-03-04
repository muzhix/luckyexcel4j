package net.hanbd.luckyexcel4j.lucky.poi.base;

import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheetView;

import java.util.List;

/**
 * Luckysheet 工作表初始化配置信息
 *
 * @author hanbd
 * @see <a
 * href="https://mengshukeji.gitee.io/LuckysheetDocs/zh/guide/sheet.html#%E5%88%9D%E5%A7%8B%E5%8C%96%E9%85%8D%E7%BD%AE">
 * 初始化配置</a>
 */
@Data
@NoArgsConstructor
public class SheetMeta {
    /**
     * 工作表名称
     */
    protected String name;
    /**
     * 工作表颜色
     */
    protected String color;
    /**
     * 工作表索引
     */
    protected Integer index;
    /**
     * 工作表激活状态, 激活表示当前显示的是该工作表(对应{@link CTSheetView#getTabSelected()}). 1,选中; 0,未选中
     */
    protected Integer status;
    /**
     * 工作表顺序
     */
    protected Integer order;
    /**
     * 行数,包括空行
     */
    protected Integer row;
    /**
     * 列数,包括空列
     */
    protected Integer column;

    protected SheetConfig config;
    /**
     * 选中的区域(支持多选)
     */
    @JsonProperty("luckysheet_select_save")
    protected List<Selection> luckySheetSelectSave;
    /**
     * 左右滚动条位置
     */
    protected Integer scrollLeft;
    /**
     * 上下滚动条位置
     */
    protected Integer scrollTop;
    /**
     * 初始化使用的单元格数据
     */
    @JsonProperty("celldata")
    protected List<Cell> cellData;

    /**
     * 图表
     */
    protected List<SheetChart> sheetChart;
    /**
     * 是否数据透视表
     */
    protected Boolean isPivotTable;
    /**
     * 数据透视表
     */
    protected SheetPivotTable pivotTable;

    /**
     * 条件格式
     */
    protected ConditionFormat conditionFormatSave;
    /**
     * TODO 待确认具体类型
     */
    protected Object freezen;
    /**
     * 公式链 TODO 待支持
     *
     * @see <a
     * href="https://mengshukeji.gitee.io/LuckysheetDocs/zh/guide/sheet.html#calcchain">calcChain</a>
     */
    protected Object calcChain;
    /**
     * 此sheet页面的缩放比例, 为0~1之间的二位小数数字。eg: 1、0.1、0.56
     */
    protected Double zoomRatio;
    /**
     * 是否显示网格线. 1, 显示; 0,隐藏 TODO 该参数是不是应该再配置项中? 与版本有关?
     */
    protected Integer showGridLines;

    /**
     * 默认列宽(pixel)
     */
    protected Integer defaultColWidth;
    /**
     * 默认行高(pixel)
     */
    protected Integer defaultRowHeight;
    /**
     * 图片
     */
    protected List<SheetImage> sheetImage;
}
