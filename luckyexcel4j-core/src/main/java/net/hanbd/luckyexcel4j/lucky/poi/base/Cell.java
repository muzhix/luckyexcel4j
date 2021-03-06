package net.hanbd.luckyexcel4j.lucky.poi.base;

import com.fasterxml.jackson.annotation.JsonIgnore;
import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.AllArgsConstructor;
import lombok.Data;
import net.hanbd.luckyexcel4j.lucky.poi.enums.BorderStyle;
import net.hanbd.luckyexcel4j.lucky.poi.enums.CellFormatType;
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
        cellValue.setValue(cell.getRawValue());
        cellValue.setCellType(lsCellFormat);
        cellValue.setBackground(cellStyle.getFillBackgroundColorColor().getARGBHex());

        // font
        XSSFFont font = cellStyle.getFont();
        cellValue.setFontsize((int) font.getFontHeight());
        cellValue.setFontColor(font.getXSSFColor().getARGBHex());
        cellValue.setFontFamily(font.getFontName());
        cellValue.setBold(font.getBold() ? 1 : 0);
        cellValue.setItalic(font.getItalic() ? 1 : 0);
        cellValue.setCancelLine(font.getStrikeout() ? 1 : 0);

        // horizontal and vertical
        int ht = 0, vt = 0;
        switch (cellStyle.getAlignment()) {
            case LEFT:
                ht = 1;
                break;
            case RIGHT:
                ht = 2;
                break;
            default:
                break;
        }
        switch (cellStyle.getVerticalAlignment()) {
            case TOP:
                vt = 1;
                break;
            case BOTTOM:
                vt = 2;
                break;
            default:
                break;
        }
        cellValue.setHorizontalType(ht);
        cellValue.setVerticalType(vt);
        cellValue.setTextBreak(cellStyle.getWrapText() ? 2 : 1);
        // TODO optimize rotation
        short rotation = cellStyle.getRotation();
        cellValue.setRotateText((int) rotation);
        cellValue.setTextRotate(0);

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
                .style(BorderStyle.valueOf(borderStyle).getStyle())
                .color(cellStyle.getBorderColor(borderSide).getARGBHex()).build();
    }


}
