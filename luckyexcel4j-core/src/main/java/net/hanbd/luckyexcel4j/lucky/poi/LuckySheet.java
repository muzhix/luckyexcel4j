package net.hanbd.luckyexcel4j.lucky.poi;

import cn.hutool.core.util.IdUtil;
import com.google.common.collect.Lists;
import lombok.Getter;
import net.hanbd.luckyexcel4j.lucky.poi.base.*;
import net.hanbd.luckyexcel4j.utils.PoiUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.model.CalculationChain;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.*;

import javax.annotation.Nullable;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;

import static net.hanbd.luckyexcel4j.utils.PojoUtil.*;

/**
 * @author hanbd
 */
public class LuckySheet {
    @Getter
    private final XSSFSheet sheet;
    private final XSSFWorkbook workbook;
    private final StylesTable stylesTable;
    private final SharedStringsTable sharedStringsTable;
    private final CalculationChain calcChain;

    private final SheetMeta sheetMeta;

    public LuckySheet(XSSFSheet sheet) {
        this.sheetMeta = new SheetMeta();

        this.sheet = sheet;
        this.workbook = this.sheet.getWorkbook();

        this.stylesTable = workbook.getStylesSource();
        this.sharedStringsTable = workbook.getSharedStringSource();
        this.calcChain = workbook.getCalculationChain();
    }

    public SheetMeta parse() {
        // sheet基础信息
        sheetMeta.setName(sheet.getSheetName());
        sheetMeta.setIndex(IdUtil.fastSimpleUUID());
        sheetMeta.setOrder(workbook.getSheetIndex(sheet));
        // 关键值初始化
        sheetMeta.setConfig(new SheetConfig());
        sheetMeta.setCellData(Lists.newArrayList());

        resolveSheetView();
        resolveTabColor();
        resolveColWidthAndRowHeightAndAddCell();
        // TODO resolve formulaRefList
        // TODO resolve calcChain (workbook.getCalculationChain())
        resolveMergeCells();

        return sheetMeta;
    }

    private void resolveSheetView() {
        CTSheetViews sheetViews = sheet.getCTWorksheet().getSheetViews();
        if (sheetViews.sizeOfSheetViewArray() <= 0) {
            return;
        }
        CTSheetView sheetView = sheetViews.getSheetViewArray(0);

        this.sheetMeta.setShowGridLines(sheetView.getShowGridLines());
        this.sheetMeta.setSelected(sheetView.getTabSelected());
        this.sheetMeta.setZoomRatio(getLongDefault(sheetView.getZoomScale(), 100L) / 100.0);
        CellAddress activeCell = sheet.getActiveCell();
        Range range = new Range(activeCell.getRow(), activeCell.getColumn());
        this.sheetMeta.setSelectRangesSave(Lists.newArrayList(range));
    }

    private void resolveTabColor() {
        @Nullable XSSFColor tabColor = sheet.getTabColor();
        if (Objects.nonNull(tabColor)) {
            sheetMeta.setColor(tabColor.getARGBHex());
        }
    }

    /**
     * 处理行列宽高,行列隐藏问题,以及单元格值
     */
    private void resolveColWidthAndRowHeightAndAddCell() {
        // resolve default
        int defaultHeightPoint = getIntDefault((int) sheet.getDefaultRowHeight(), 19);
        sheetMeta.setDefaultColWidth(
                PoiUtil.getColumnWidthPixel((double) sheet.getDefaultColumnWidth()));
        sheetMeta.setDefaultRowHeight(PoiUtil.getRowHeightPixel((double) defaultHeightPoint));

        SheetConfig sheetConfig = sheetMeta.getConfig();
        // resolve column information
        List<CTCols> colsList = sheet.getCTWorksheet().getColsList();
        for (CTCols cols : colsList) {
            for (CTCol col : cols.getColList()) {
                if (col.getMin() == 0L || col.getMax() == 0L) {
                    continue;
                }
                // ECMA规范中列是1 based, luckysheet中列是0 based,所有此处要减1
                int minNum = (int) col.getMin() - 1, maxNum = (int) col.getMax() - 1;
                @Nullable Double widthNum = getDoubleDefault(col.getWidth(), null);
                boolean hidden = col.getHidden();

                for (int i = minNum; i <= maxNum; i++) {
                    // columnlen
                    if (Objects.nonNull(widthNum)) {
                        // 添加列宽
                        sheetConfig.getColumnWidth().put(i, PoiUtil.getColumnWidthPixel(widthNum));
                    }

                    // colhidden
                    if (hidden) {
                        // 隐藏列
                        sheetConfig.getHiddenColumns().add(i);
                        // 删除列宽
                        sheetConfig.getColumnWidth().remove(i);
                    }
                }
            }
        }

        // resolve rows and cells
        List<CTRow> rowList = sheet.getCTWorksheet().getSheetData().getRowList();
        Map<Integer, Integer> rowHeightMap = sheetConfig.getRowHeight();
        Set<Integer> hiddenRows = sheetConfig.getHiddenRows();
        for (Row row : sheet) {
            // 此处拿到的行值是0 based
            int rowNo = row.getRowNum();
            // rowlen
            rowHeightMap.put(rowNo, PoiUtil.getRowHeightPixel((double) row.getHeight()));
            // rowhidden
            if (row.getZeroHeight()) {
                // 隐藏行
                hiddenRows.add(rowNo);
                // 删除行高
                sheetConfig.getRowHeight().remove(rowNo);
            }
            // cells
            for (org.apache.poi.ss.usermodel.Cell cell : row) {
                Cell lsCell = new Cell((XSSFCell) cell);
                if (Objects.nonNull(lsCell.getCellBorder())) {
                    sheetConfig.getBorderInfos().add(lsCell.getCellBorder());
                }
                sheetMeta.getCellData().add(lsCell);
            }
        }
    }

    /**
     * 处理合并单元格
     */
    private void resolveMergeCells() {
        Map<String, MergeCell> mergeMap = sheetMeta.getConfig().getMerge();
        List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
        for (CellRangeAddress cellRange : mergedRegions) {
            MergeCell merge = MergeCell.builder()
                    .startCol(cellRange.getFirstColumn())
                    .startRow(cellRange.getFirstRow())
                    .colsNum(cellRange.getLastColumn() - cellRange.getFirstColumn() + 1)
                    .rowsNum(cellRange.getLastRow() - cellRange.getFirstRow() + 1).build();
            mergeMap.put(merge.mainCellStr(), merge);
        }
    }
}
