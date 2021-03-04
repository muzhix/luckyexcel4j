package net.hanbd.luckyexcel4j.lucky.poi.base;

import com.fasterxml.jackson.annotation.JsonProperty;

import java.util.List;

/**
 * luckysheet 条件格式配置信息
 *
 * @author hanbd
 * @see <a
 * href="https://mengshukeji.gitee.io/LuckysheetDocs/zh/guide/sheet.html#luckysheet-conditionformat-save">
 * luckysheet-conditionformat-save</a>
 */
public class ConditionFormat {
    /**
     * 突出显示单元格规则和项目选区规则
     *
     * <p>Option:default,dataBar,colorGradation,icons TODO 使用enum替换string
     */
    private String type;
    /**
     * 条件应用范围
     */
    @JsonProperty("cellrange")
    private Selection cellRange;
    /**
     * 格式,根据不同的{@link ConditionFormat#type},实际对象类型不同
     */
    private Object format;
    /**
     * 类型
     *
     * <p>Detailed settings,comparison parameters
     */
    private String conditionName;
    /**
     * 条件值所在单元格
     *
     * <p>Detailed settings,comparison range
     */
    private List<Selection> conditionRange;
    /**
     * 条件值 TODO 具体类型待确认
     */
    private List<Object> conditionValue;
}
