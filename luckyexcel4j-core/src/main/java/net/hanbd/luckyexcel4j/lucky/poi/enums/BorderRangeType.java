package net.hanbd.luckyexcel4j.lucky.poi.enums;

import lombok.AllArgsConstructor;
import lombok.Getter;

/**
 * 边框范围类型
 *
 * @author hanbd
 * @see <a href="https://mengshukeji.gitee.io/LuckysheetDocs/zh/guide/sheet.html#config-borderinfo">border=info</a>
 */
@AllArgsConstructor
public enum BorderRangeType {
    /**
     * 区域
     */
    RANGE("range"),
    /**
     * 单个单元格
     */
    CELL("cell");
    @Getter
    private final String type;
}
