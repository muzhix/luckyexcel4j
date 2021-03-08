package net.hanbd.luckyexcel4j.utils;

import com.google.common.primitives.Ints;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.awt.*;

class PoiUtilTest {
    XSSFColor notnullRgb;
    XSSFColor nullRgb;
    XSSFColor nullColor;

    @BeforeEach
    void setUp() {
        Color red = Color.RED;
        byte[] bytes = Ints.toByteArray(red.getRGB());
        notnullRgb = new XSSFColor();
        notnullRgb.setRGB(bytes);

        nullRgb = new XSSFColor();
        nullRgb.setRGB(null);

        nullColor = null;
    }

    @AfterEach
    void tearDown() {
    }

    @Test
    void getRgbHexStr() {
        Assertions.assertEquals("#ff0000", PoiUtil.getRgbHexStr(notnullRgb));
        Assertions.assertNull(PoiUtil.getRgbHexStr(nullRgb));
        Assertions.assertNull(PoiUtil.getRgbHexStr(nullColor));
    }

    @Test
    void getRgbStr() {
        Assertions.assertEquals("rgb(255,0,0)", PoiUtil.getRgbStr(notnullRgb));
        Assertions.assertNull(PoiUtil.getRgbStr(nullRgb));
        Assertions.assertNull(PoiUtil.getRgbStr(nullColor));
    }
}