package net.hanbd.luckyexcel4j.utils;

import javax.annotation.Nonnull;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.Date;

/**
 * @author hanbd
 */
public class DateUtil {
    /**
     * {@link Date} -> {@link LocalDateTime}
     *
     * @param date
     * @return
     */
    public static LocalDateTime transformForm(@Nonnull Date date) {
        return date.toInstant().atZone(ZoneId.systemDefault()).toLocalDateTime();
    }

    /**
     * {@link LocalDateTime} -> {@link Date}
     *
     * @param localDateTime
     * @return
     */
    public static Date transformForm(@Nonnull LocalDateTime localDateTime) {
        return Date.from(localDateTime.atZone(ZoneId.systemDefault()).toInstant());
    }
}
