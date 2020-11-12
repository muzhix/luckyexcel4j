package net.hanbd.luckyexcel4j.lucky.poi;

import lombok.Getter;
import net.hanbd.luckyexcel4j.lucky.poi.base.LuckyFileBase;
import net.hanbd.luckyexcel4j.lucky.poi.base.LuckyFileInfo;
import net.hanbd.luckyexcel4j.lucky.poi.base.LuckySheetBase;
import net.hanbd.luckyexcel4j.utils.DateUtil;
import org.apache.poi.ooxml.POIXMLProperties.CoreProperties;
import org.apache.poi.ooxml.POIXMLProperties.ExtendedProperties;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * @author hanbd
 */
public class LuckyFile {
    @Getter
    private final XSSFWorkbook excel;
    /**
     * 源文件
     */
    @Getter
    private final File file;

    public LuckyFile(File file, PackageAccess access) {
        try {
            this.file = file;
            this.excel = new XSSFWorkbook(OPCPackage.open(file, access));
        } catch (IOException | OpenXML4JException e) {
            throw new RuntimeException("加载文件失败", e);
        }
    }

    /**
     * xlsx file -> {@link LuckyFileBase}
     *
     * @return LuckyFileBase 用于前端渲染luckysheet表格
     */
    public LuckyFileBase parse() {
        return LuckyFileBase.builder().info(this.getWorkbookInfo()).sheets(this.getSheets()).build();
    }

    /**
     * 获得excel工作表的基本信息
     *
     * @return
     */
    private LuckyFileInfo getWorkbookInfo() {
        CoreProperties coreProperties = excel.getProperties().getCoreProperties();
        ExtendedProperties extProperties = excel.getProperties().getExtendedProperties();
        return LuckyFileInfo.builder()
                .name(file.getName())
                .appVersion(extProperties.getAppVersion())
                .company(extProperties.getCompany())
                .creator(coreProperties.getCreator())
                .lastModifiedBy(coreProperties.getLastModifiedByUser())
                .createdTime(DateUtil.transformForm(coreProperties.getCreated()))
                .modifiedTime(DateUtil.transformForm(coreProperties.getModified()))
                .build();
    }

    private List<LuckySheetBase> getSheets() {
        int sheetNumbers = this.excel.getNumberOfSheets();

        List<LuckySheetBase> sheets = new ArrayList<>(sheetNumbers);
        for (int i = 0; i < sheetNumbers; i++) {
            XSSFSheet sheet = this.excel.getSheetAt(i);
            LuckySheet luckySheet = new LuckySheet(sheet);
            sheets.add(luckySheet.getSheetBase());
        }

        return sheets;
    }
}
