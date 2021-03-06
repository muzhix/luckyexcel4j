package net.hanbd.luckyexcel4j.lucky.poi.base;

import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * 批注
 *
 * @author hanbd
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class Comment {
    /**
     * 批注框左边距
     */
    private Integer left;
    /**
     * 批注框上边距
     */
    private Integer top;
    /**
     * 批注框宽度
     */
    private Integer width;
    /**
     * 批注框高度
     */
    private Integer height;
    /**
     * 批注内容
     */
    private String value;
    /**
     * 批注框是否显示
     */
    @JsonProperty("isshow")
    private Boolean show;
}
