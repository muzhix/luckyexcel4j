package net.hanbd.luckyexcel4j.lucky.poi.base;

import com.fasterxml.jackson.annotation.JsonIgnore;
import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import net.hanbd.luckyexcel4j.lucky.poi.enums.*;
import net.hanbd.luckyexcel4j.utils.PoiUtil;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCell;

import javax.validation.constraints.Min;
import java.text.DecimalFormat;
import java.util.Objects;

/**
 * @author hanbd
 * @see <a
 * href="https://mengshukeji.gitee.io/luckysheetdocs/zh/guide/cell.html#%E5%8D%95%E5%85%83%E6%A0%BC">
 * 单元格</a>
 */
@Slf4j
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
    @JsonIgnore
    private final CTCell ctCell;
    @JsonIgnore
    private final XSSFWorkbook workbook;
    @JsonIgnore
    private CellBorder cellBorder;

    public Cell(XSSFCell cell) {
        this.cell = cell;
        this.ctCell = cell.getCTCell();
        this.workbook = cell.getSheet().getWorkbook();

        this.row = cell.getRowIndex();
        this.column = cell.getColumnIndex();
        log.debug("row: {}, col: {}", row, column);
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
        setValueAndMonitor(cellValue);
        cellValue.setCellType(lsCellFormat);
        String bgColorHex = PoiUtil.getRgbHexStr(cellStyle.getFillBackgroundColorColor());
        if (Objects.nonNull(bgColorHex)) {
            cellValue.setBackground(bgColorHex);
        }

        // font
        XSSFFont font = cellStyle.getFont();
        cellValue.setFontsize(font.getFontHeightInPoints());
        String fontColorHex = PoiUtil.getRgbHexStr(font.getXSSFColor());
        if (Objects.nonNull(fontColorHex)) {
            cellValue.setFontColor(fontColorHex);
        }
        log.debug("{} fontName: {}", cell.getAddress(), font.getFontName());
        cellValue.setFontFamily(FontFamily.of(font.getFontName()));
        cellValue.setBold(font.getBold());
        cellValue.setItalic(font.getItalic());
        cellValue.setCancelLine(font.getStrikeout());

        // horizontal and vertical
        cellValue.setHorizontalType(CellHorizontalType.of(cellStyle.getAlignment(), cell.getCellType()));
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
                borderValue.setRight(genBorderStyle(cellStyle, borderRight, XSSFCellBorder.BorderSide.RIGHT));
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
        String bdColorStr = PoiUtil.getRgbStr(cellStyle.getBorderColor(borderSide));
        CellBorder.Style style = new CellBorder.Style();
        style.setStyle(BorderStyle.of(borderStyle).getStyle());
        if (Objects.nonNull(bdColorStr)) {
            style.setColor(bdColorStr);
        }
        return style;
    }

    private static final DecimalFormat NUMERIC_FORMAT = new DecimalFormat("#.######");

    public void setValueAndMonitor(CellValue cellValue) {
        if (Objects.isNull(cell)) {
            return;
        }
        String strVal = "";
        switch (cell.getCellType()) {
            case STRING:
                strVal = cell.getStringCellValue();
                break;
            case NUMERIC:
                strVal = NUMERIC_FORMAT.format(cell.getNumericCellValue());
                break;
            case BOOLEAN:
                strVal = String.valueOf(cell.getBooleanCellValue());
                break;
            case FORMULA:
                try {
                    strVal = NUMERIC_FORMAT.format(cell.getNumericCellValue());
                } catch (IllegalStateException e) {
                    if (e.getMessage().contains("from a STRING cell")) {
                        strVal = cell.getStringCellValue();
                    } else if (e.getMessage().contains("from a ERROR formula cell")) {
                        strVal = String.valueOf(cell.getErrorCellValue());
                    }
                } catch (Exception e) {
                    log.error("获取FORMULA单元格内容出错, error: {}", e.getMessage());
                    e.printStackTrace();
                    strVal = cell.toString();
                }
            default:
                break;
        }
        String monitor = strVal;

        cellValue.setValue(strVal);
        cellValue.setMonitor(monitor);
    }
}

