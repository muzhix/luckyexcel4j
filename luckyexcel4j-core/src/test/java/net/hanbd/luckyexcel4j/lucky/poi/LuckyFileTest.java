package net.hanbd.luckyexcel4j.lucky.poi;

import cn.hutool.core.io.FileUtil;
import net.hanbd.luckyexcel4j.lucky.poi.base.FileMeta;
import net.hanbd.luckyexcel4j.utils.JsonUtil;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.nio.charset.StandardCharsets;

class LuckyFileTest {

    File excelFile = new File("F:\\tmp\\officexml\\test.xlsx");
    File jsonFile = new File("F:\\tmp\\officexml\\test.json");

    @BeforeEach
    void setUp() {
    }

    @AfterEach
    void tearDown() {
    }

    @Test
    void parse() {
        LuckyFile luckyFile = new LuckyFile(excelFile, PackageAccess.READ);
        FileMeta fileMeta = luckyFile.parse();
        String json = JsonUtil.stringify(fileMeta);
        FileUtil.writeString(json, jsonFile, StandardCharsets.UTF_8);
    }
}