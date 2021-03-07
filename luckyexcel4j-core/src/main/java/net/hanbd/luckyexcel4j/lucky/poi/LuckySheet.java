package net.hanbd.luckyexcel4j.lucky.poi;

import cn.hutool.core.util.IdUtil;
import com.google.common.collect.Lists;
import lombok.Getter;
import lombok.extern.slf4j.Slf4j;
import net.hanbd.luckyexcel4j.lucky.poi.base.*;
import net.hanbd.luckyexcel4j.utils.PoiUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.model.CalculationChain;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCell;
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
@Slf4j
public class LuckySheet {
    @Getter
    private final XSSFSheet sheet;
    private final CTWorksheet ctSheet;
    private final XSSFWorkbook workbook;
    private final StylesTable stylesTable;
    private final SharedStringsTable sharedStringsTable;
    private final CalculationChain calcChain;

    private final SheetMeta sheetMeta;

    public LuckySheet(XSSFSheet sheet) {
        this.sheetMeta = new SheetMeta();

        this.sheet = sheet;
        this.workbook = this.sheet.getWorkbook();
        ctSheet = sheet.getCTWorksheet();

        this.stylesTable = workbook.getStylesSource();
        this.sharedStringsTable = workbook.getSharedStringSource();
        this.calcChain = workbook.getCalculationChain();
    }

    public SheetMeta parse() {
        // sheet基础信息
        sheetMeta.setName(sheet.getSheetName());
        sheetMeta.setIndex(IdUtil.fastSimpleUUID());
        int sheetIdx = workbook.getSheetIndex(sheet);
        sheetMeta.setOrder(sheetIdx);
        sheetMeta.setHidden(workbook.isSheetHidden(sheetIdx));

        resolveSheetView();
        resolveTabColor();
        resolveMergeCells();
        resolveColWidthAndRowHeightAndAddCell();
        // TODO resolve formulaRefList
        // TODO resolve calcChain (workbook.getCalculationChain())
        // TODO resolve pivotTable

        return sheetMeta;
    }

    private void resolveSheetView() {
        // apache中没有获取zoomScale的方法,这里使用更底层的api来获取
        CTSheetView sheetView = getDefaultSheetView();
        double zoomRatio = Objects.isNull(sheetView) ?
                1.0 : getLongDefault(sheetView.getZoomScale(), 100L) / 100.0;
        this.sheetMeta.setZoomRatio(zoomRatio);
        this.sheetMeta.setShowGridLines(sheet.isDisplayGridlines());
        this.sheetMeta.setSelected(sheet.isSelected());
        CellAddress activeCell = sheet.getActiveCell();
        Range range = new Range(activeCell.getRow(), activeCell.getColumn());
        this.sheetMeta.setSelectRangesSave(Lists.newArrayList(range));
    }

    @Nullable
    private CTSheetView getDefaultSheetView() {
        if (!ctSheet.isSetSheetViews()) {
            return null;
        }
        final CTSheetViews views = ctSheet.getSheetViews();
        int sz = views.sizeOfSheetViewArray();
        if (sz == 0) {
            return null;
        }
        return views.getSheetViewArray(sz - 1);
    }

    private void resolveTabColor() {
        String tabColorHex = PoiUtil.getRgbHexStr(sheet.getTabColor());
        if (Objects.nonNull(tabColorHex)) {
            sheetMeta.setColor(tabColorHex);
        }
    }

    /**
     * 处理行列宽高,行列隐藏问题,以及单元格值
     */
    private void resolveColWidthAndRowHeightAndAddCell() {
        // resolve default
        int defaultHeightInPoint = getIntDefault((int) sheet.getDefaultRowHeightInPoints(), 19);
        Integer defaultRowHeight = PoiUtil.getRowHeightPixel((double) defaultHeightInPoint);
        sheetMeta.setDefaultColWidth(PoiUtil.getColumnWidthPixel((double) sheet.getDefaultColumnWidth()));
        sheetMeta.setDefaultRowHeight(defaultRowHeight);

        SheetConfig sheetConfig = sheetMeta.getConfig();
        // resolve column information
        Map<Integer, Integer> columnWidthMap = sheetConfig.getColumnWidth();
        Set<Integer> hiddenColumns = sheetConfig.getHiddenColumns();
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
                        columnWidthMap.put(i, PoiUtil.getColumnWidthPixel(widthNum));
                    }

                    // colhidden
                    if (hidden) {
                        // 隐藏列
                        hiddenColumns.add(i);
                        // 删除列宽
//                        columnWidthMap.remove(i);
                    }
                }
            }
        }

        // resolve rows and cells
        List<CTRow> rowList = sheet.getCTWorksheet().getSheetData().getRowList();
        Map<Integer, Integer> rowHeightMap = sheetConfig.getRowHeight();
        Set<Integer> hiddenRows = sheetConfig.getHiddenRows();
        int colNum = 0;
        for (Row row : sheet) {
            int rowNo = row.getRowNum();

            int rh = PoiUtil.getRowHeightPixel((double) row.getHeightInPoints());
            if (!Objects.equals(defaultRowHeight, rh)) {
                rowHeightMap.put(rowNo, rh);
            }
            if (row.getZeroHeight()) {
                hiddenRows.add(rowNo);
//                rowHeightMap.remove(rowNo);
            }

            // cells
            List<Border> borderInfos = sheetConfig.getBorderInfos();
            for (org.apache.poi.ss.usermodel.Cell cell : row) {
                Cell lsCell = new Cell((XSSFCell) cell);
                if (Objects.nonNull(lsCell.getCellBorder())) {
                    borderInfos.add(lsCell.getCellBorder());
                }
                sheetMeta.getCellData().add(lsCell);
            }

            colNum = Math.max(colNum, row.getLastCellNum());
        }
        // 设定行数和列数
        sheetMeta.setRow(sheet.getLastRowNum() + 1);
        sheetMeta.setColumn(colNum);
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
