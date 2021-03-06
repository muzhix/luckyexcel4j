package net.hanbd.luckyexcel4j.lucky.poi.base;

import lombok.Data;
import lombok.NoArgsConstructor;
import net.hanbd.luckyexcel4j.lucky.poi.enums.BorderRangeType;

/**
 * @author hanbd
 */
@Data
@NoArgsConstructor
public class Border {
    /**
     * 边框范围类型
     */
    protected BorderRangeType rangeType;
}
