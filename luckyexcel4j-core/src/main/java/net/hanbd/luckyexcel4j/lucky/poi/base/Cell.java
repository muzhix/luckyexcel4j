package net.hanbd.luckyexcel4j.lucky.poi.base;

import com.fasterxml.jackson.annotation.JsonIgnore;
import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.AllArgsConstructor;
import lombok.Data;
import net.hanbd.luckyexcel4j.lucky.poi.enums.*;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCell;

import javax.validation.constraints.Min;

/**
 * @author hanbd
 * @see <a
 * href="https://mengshukeji.gitee.io/luckysheetdocs/zh/guide/cell.html#%E5%8D%95%E5%85%83%E6%A0%BC">
 * 单元格</a>
 */
@Data
@AllArgsConstructor
public class Cell {
    /**
     * 行index
     */
    @JsonProperty("r")
    @Min(0)
    private Integer row;
    /**
     * 列index
     */
    @JsonProperty("c")
    @Min(0)
    private Integer column;
    /**
     * 单元格值
     */
    @JsonProperty("v")
    private CellValue value;

    @JsonIgnore
    private final XSSFCell cell;
    private final CTCell ctCell;
    private final XSSFWorkbook workbook;
    @JsonIgnore
    private CellBorder cellBorder;

    public Cell(XSSFCell cell) {
        this.cell = cell;
        this.ctCell = cell.getCTCell();
        this.workbook = cell.getSheet().getWorkbook();

        this.row = cell.getRowIndex();
        this.column = cell.getColumnIndex();
        this.value = generateValue();
    }

    private CellValue generateValue() {
        SharedStringsTable sst = workbook.getSharedStringSource();
        CellValue cellValue = new CellValue();

        /*
         * style
         */
        XSSFCellStyle cellStyle = cell.getCellStyle();

        CellFormat lsCellFormat = CellFormat.builder()
                .format(cellStyle.getDataFormatString())
                .type(CellFormatType.of(cell.getCellType())).build();
        // TODO 原始值和显示值的赋值待优化
        cellValue.setValue(cell.getRawValue());
        cellValue.setMonitor(cell.getRawValue());
        cellValue.setCellType(lsCellFormat);
        cellValue.setBackground(cellStyle.getFillBackgroundColorColor().getARGBHex());

        // font
        XSSFFont font = cellStyle.getFont();
        cellValue.setFontsize(font.getFontHeightInPoints());
        cellValue.setFontColor(font.getXSSFColor().getARGBHex());
        cellValue.setFontFamily(FontFamily.of(font.getFontName()));
        cellValue.setBold(font.getBold());
        cellValue.setItalic(font.getItalic());
        cellValue.setCancelLine(font.getStrikeout());

        // horizontal and vertical
        cellValue.setHorizontalType(CellHorizontalType.of(cellStyle.getAlignment()));
        cellValue.setVerticalType(CellVerticalType.of(cellStyle.getVerticalAlignment()));
        cellValue.setTextBreak(cellStyle.getWrapText() ? TextBreakType.LINE_WRAP : TextBreakType.OVERFLOW);
        short rotation = cellStyle.getRotation();
        cellValue.setRotateText(rotation);

        // border
        if (cellStyle.getCoreXf().getApplyBorder()) {
            cellBorder = new CellBorder();
            CellBorder.Value borderValue = cellBorder.getValue();
            borderValue.setRow(this.row);
            borderValue.setCol(this.column);

            org.apache.poi.ss.usermodel.BorderStyle borderTop = cellStyle.getBorderTop();
            org.apache.poi.ss.usermodel.BorderStyle borderBottom = cellStyle.getBorderBottom();
            org.apache.poi.ss.usermodel.BorderStyle borderLeft = cellStyle.getBorderLeft();
            org.apache.poi.ss.usermodel.BorderStyle borderRight = cellStyle.getBorderRight();

            if (borderTop != org.apache.poi.ss.usermodel.BorderStyle.NONE) {
                borderValue.setTop(genBorderStyle(cellStyle, borderTop, XSSFCellBorder.BorderSide.TOP));
            }
            if (borderBottom != org.apache.poi.ss.usermodel.BorderStyle.NONE) {
                borderValue.setBottom(genBorderStyle(cellStyle, borderBottom, XSSFCellBorder.BorderSide.BOTTOM));
            }
            if (borderLeft != org.apache.poi.ss.usermodel.BorderStyle.NONE) {
                borderValue.setLeft(genBorderStyle(cellStyle, borderLeft, XSSFCellBorder.BorderSide.LEFT));
            }
            if (borderRight != org.apache.poi.ss.usermodel.BorderStyle.NONE) {
                borderValue.setTop(genBorderStyle(cellStyle, borderRight, XSSFCellBorder.BorderSide.RIGHT));
            }
        }


        // TODO support formula
        /*if (Objects.nonNull(cell.getCellFormula())) {
            CTCellFormula f = ctCell.getF();
            //noinspection AlibabaUndefineMagicConstant
            if (f.getT().equals(STCellFormulaType.Enum.forString("shared"))) {
            }
        }*/
        return cellValue;
    }

    private CellBorder.Style genBorderStyle(XSSFCellStyle cellStyle,
                                            org.apache.poi.ss.usermodel.BorderStyle borderStyle,
                                            XSSFCellBorder.BorderSide borderSide) {
        return CellBorder.Style.builder()
                .style(BorderStyle.of(borderStyle).getStyle())
                .color(cellStyle.getBorderColor(borderSide).getARGBHex()).build();
    }


}
