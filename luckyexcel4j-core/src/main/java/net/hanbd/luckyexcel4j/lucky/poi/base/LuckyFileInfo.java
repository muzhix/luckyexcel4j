package net.hanbd.luckyexcel4j.lucky.poi.base;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.time.LocalDateTime;

/**
 * 文件基本信息
 *
 * @author hanbd
 */
@Data
@Builder
@AllArgsConstructor
@NoArgsConstructor
public class LuckyFileInfo {
    /**
     * 文件名
     */
    private String name;
    /**
     * 创建人
     */
    private String creator;
    /**
     * 最新修改人
     */
    private String lastModifiedBy;
    /**
     * 创建时间
     */
    private LocalDateTime createdTime;
    /**
     * 最新修改时间
     */
    private LocalDateTime modifiedTime;

    private String company;
    /**
     * excel version
     */
    private String appVersion;
}