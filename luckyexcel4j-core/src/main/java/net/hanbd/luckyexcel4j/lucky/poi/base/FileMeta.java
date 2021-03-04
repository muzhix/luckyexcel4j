package net.hanbd.luckyexcel4j.lucky.poi.base;

import lombok.Builder;
import lombok.Data;

import java.util.List;

/**
 * 文件
 *
 * @author hanbd
 */
@Data
@Builder
public class FileMeta {
    private ExcelFileInfo info;
    private List<SheetMeta> sheets;
}
