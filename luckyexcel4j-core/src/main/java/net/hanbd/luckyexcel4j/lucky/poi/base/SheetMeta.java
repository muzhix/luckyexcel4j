package net.hanbd.luckyexcel4j.lucky.poi.base;

import com.fasterxml.jackson.annotation.JsonFormat;
import com.fasterxml.jackson.annotation.JsonProperty;
import com.google.common.collect.Lists;
import lombok.Data;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheetView;

import javax.validation.constraints.Max;
import javax.validation.constraints.Min;
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
public class SheetMeta {
    /**
     * new 对象,并对关键值初始化
     */
    public SheetMeta() {
        this.config = new SheetConfig();
        this.cellData = Lists.newArrayList();
    }

    /**
     * 工作表名称
     */
    protected String name;
    /**
     * 工作表颜色。hex值, eg: #f20e0e
     */
    protected String color;
    /**
     * 工作表索引,作为唯一key存在。注意与{@link SheetMeta#order}做区分
     */
    protected String index;
    /**
     * 工作表激活状态, 激活表示当前选中并显示的是该工作表(对应{@link CTSheetView#getTabSelected()}). 1,选中; 0,未选中
     */
    @JsonProperty("status")
    @JsonFormat(shape = JsonFormat.Shape.NUMBER)
    protected Boolean selected;
    /**
     * 工作表顺序下标,代表工作表在底部sheet栏展示的顺序, <b>0 based</b>
     */
    protected Integer order;
    /**
     * 是否隐藏. 0 不隐藏; 1 隐藏
     */
    @JsonProperty("hide")
    @JsonFormat(shape = JsonFormat.Shape.NUMBER)
    protected Boolean hidden;
    /**
     * 行数,包括空行
     */
    protected Integer row;
    /**
     * 列数,包括空列
     */
    protected Integer column;
    /**
     * 表格行高、列宽、合并单元格、边框、隐藏行等设置
     * <p>
     * 如果为空,只能为空对象,不能为null
     */
    protected SheetConfig config;
    /**
     * 选中的区域(支持多选)
     */
    @JsonProperty("luckysheet_select_save")
    protected List<Range> selectRangesSave;
    /**
     * 左右滚动条位置
     * <p>
     * ECMA-376中无对应属性
     */
    protected Integer scrollLeft;
    /**
     * 上下滚动条位置
     * <p>
     * ECMA-376中无对应属性
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
    @JsonProperty("chart")
    protected List<SheetChart> charts;
    /**
     * 是否数据透视表
     */
    protected Boolean isPivotTable;
    /**
     * 数据透视表
     */
    protected SheetPivotTable pivotTable;

    /**
     * 筛选范围。一个选区，一个sheet只有一个筛选范围，类似luckysheet_select_save。
     * 如果仅仅只是创建一个选区打开筛选功能，则配置这个范围即可，如果还需要进一步设置详细的筛选条件，
     * 则需要另外配置同级的 {@link SheetMeta#filter} 属性
     */
    @JsonProperty("filter_select")
    protected Range filterSelect;

    /**
     * 筛选的具体设置，跟{@link SheetMeta#filterSelect}筛选范围是互相搭配的
     * TODO
     */
    protected Object filter;

    /**
     * 条件格式
     * TODO
     */
    @JsonProperty("luckysheet_conditionformat_save")
    protected List<ConditionFormat> conditionFormatSave;
    /**
     * 冻结行列设置 TODO 待确认具体类型
     */
    protected List<Object> frozen;
    /**
     * 公式链 TODO 待支持
     *
     * @see <a
     * href="https://mengshukeji.gitee.io/LuckysheetDocs/zh/guide/sheet.html#calcchain">calcChain</a>
     */
    protected Object calcChain;
    /**
     * 此sheet页面的缩放比例, 为0~1之间的二位小数数字. eg: 1、0.1、0.56.  default 1
     * <p>
     * TODO apache poi默认的 zoomScale 是 10% ~ 400%, 该值的具体赋值范围需要再详细测试
     */
    @Min(0)
    @Max(1)
    protected Double zoomRatio;
    /**
     * 是否显示网格线. 1, 显示; 0,隐藏. default true
     */
    @JsonFormat(shape = JsonFormat.Shape.NUMBER)
    protected Boolean showGridLines;
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
    @JsonProperty("image")
    protected List<SheetImage> images;
    /**
     * 数据验证的配置信息。
     * TODO
     */
    protected Object dataVerification;
}
