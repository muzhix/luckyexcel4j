package net.hanbd.luckyexcel4j.lucky.poi;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
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

import java.util.List;
import java.util.Map;
import java.util.Objects;

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

    @Getter
    private SheetMeta sheetMeta;

    public LuckySheet(XSSFSheet sheet) {
        this.sheetMeta = new SheetMeta();

        this.sheet = sheet;
        this.workbook = this.sheet.getWorkbook();

        this.stylesTable = workbook.getStylesSource();
        this.sharedStringsTable = workbook.getSharedStringSource();
        this.calcChain = workbook.getCalculationChain();
    }

    private void parse() {
        sheetMeta.setName(sheet.getSheetName());
        sheetMeta.setIndex(workbook.getSheetIndex(sheet));
        sheetMeta.setOrder(sheetMeta.getIndex());
        sheetMeta.setConfig(new SheetConfig());
        sheetMeta.setCellData(Lists.newArrayList());

        resolveSheetView();
        resolveTabColor();
        resolveColWidthAndRowHeightAndAddCell();
        // TODO resolve formulaRefList
        // TODO resolve calcChain (workbook.getCalculationChain())
        resolveMergeCells();
    }

    private void resolveSheetView() {
        CTSheetViews sheetViews = sheet.getCTWorksheet().getSheetViews();
        if (sheetViews.sizeOfSheetViewArray() <= 0) {
            return;
        }
        CTSheetView sheetView = sheetViews.getSheetViewArray(0);

        this.sheetMeta.setShowGridLines(sheetView.getShowGridLines() ? 1 : 0);
        this.sheetMeta.setStatus(sheetView.getTabSelected() ? 1 : 0);
        this.sheetMeta.setZoomRatio(getLongDefault(sheetView.getZoomScale(), 100L) / 100.0);
        CellAddress activeCell = sheet.getActiveCell();
        Range range =
                new Range(activeCell.getRow(), activeCell.getColumn());
        this.sheetMeta.setSelectRangesSave(Lists.newArrayList(range));
    }

    private void resolveTabColor() {
        XSSFColor tabColor = sheet.getTabColor();
        sheetMeta.setColor(tabColor.getARGBHex());
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

        // resolve cols
        List<CTCols> colsList = sheet.getCTWorksheet().getColsList();
        for (CTCols cols : colsList) {
            for (CTCol col : cols.getColList()) {
                if (col.getMin() == 0L || col.getMax() == 0L) {
                    continue;
                }
                int minNum = (int) col.getMin() - 1, maxNum = (int) col.getMax() - 1;
                Double widthNum = getDoubleDefault(col.getWidth(), null);
                boolean hidden = col.getHidden();

                for (int i = minNum; i <= maxNum; i++) {
                    // columnlen
                    if (Objects.nonNull(widthNum)) {
                        if (Objects.isNull(sheetMeta.getConfig().getColumnWidth())) {
                            sheetMeta.getConfig().setColumnWidth(Maps.newHashMap());
                        }
                        // 添加列宽
                        sheetMeta.getConfig().getColumnWidth().put(i, PoiUtil.getColumnWidthPixel(widthNum));
                    }

                    // colhidden
                    if (hidden) {
                        if (Objects.isNull(sheetMeta.getConfig().getColumnHidden())) {
                            sheetMeta.getConfig().setColumnHidden(Maps.newHashMap());
                        }
                        // 隐藏列
                        sheetMeta.getConfig().getColumnHidden().put(i, 0);
                        // 删除列宽
                        if (Objects.nonNull(sheetMeta.getConfig().getColumnWidth())) {
                            sheetMeta.getConfig().getColumnWidth().remove(i);
                        }
                    }
                }
            }
        }

        // resolve rows and cells
        List<CTRow> rowList = sheet.getCTWorksheet().getSheetData().getRowList();
        if (Objects.isNull(sheetMeta.getConfig().getRowHeight())) {
            sheetMeta.getConfig().setRowHeight(Maps.newHashMap());
        }
        if (Objects.isNull(sheetMeta.getConfig().getRowHidden())) {
            sheetMeta.getConfig().setRowHidden(Maps.newHashMap());
        }
        Map<Integer, Integer> rowHeightMap = sheetMeta.getConfig().getRowHeight();
        Map<Integer, Integer> rowHiddenMap = sheetMeta.getConfig().getRowHidden();
        for (Row row : sheet) {
            int rowNo = row.getRowNum();
            // rowlen
            rowHeightMap.put(rowNo, PoiUtil.getRowHeightPixel((double) row.getHeight()));
            // rowhidden
            if (row.getZeroHeight()) {
                // 隐藏行
                rowHiddenMap.put(rowNo, 0);
                // 删除行高
                sheetMeta.getConfig().getRowHeight().remove(rowNo);
            }
            // cells
            for (org.apache.poi.ss.usermodel.Cell cell : row) {
                Cell lsCell = new Cell((XSSFCell) cell);
                if (Objects.nonNull(lsCell.getCellBorder())) {
                    if (Objects.isNull(this.sheetMeta.getConfig().getBorderInfo())) {
                        this.sheetMeta.getConfig().setBorderInfo(Lists.newArrayList());
                    }
                    this.sheetMeta.getConfig().getBorderInfo().add(lsCell.getCellBorder());
                }
                this.sheetMeta.getCellData().add(lsCell);
            }
        }
    }

    /**
     * 处理合并单元格
     */
    private void resolveMergeCells() {
        if (Objects.isNull(sheetMeta.getConfig().getMerge())) {
            sheetMeta.getConfig().setMerge(Maps.newHashMap());
        }
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
