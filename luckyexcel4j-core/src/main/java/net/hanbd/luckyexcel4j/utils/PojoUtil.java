package net.hanbd.luckyexcel4j.utils;

import cn.hutool.core.util.StrUtil;

import java.util.Objects;
import java.util.function.Predicate;

/**
 * @author hanbd
 */
public class PojoUtil {

    /**
     * 如果provide为{@code null}或空字符串,则返回defaultVal
     *
     * @param provide
     * @param defaultVal
     * @return
     */
    public static String getStringDefault(String provide, String defaultVal) {
        return StrUtil.isBlank(provide) ? defaultVal : provide;
    }

    /**
     * 如果provide为{@code null}或{@code 0L},则返回defaultVal
     *
     * @param provide
     * @param defaultVal
     * @return
     */
    public static Long getLongDefault(Long provide, Long defaultVal) {
        return Objects.isNull(provide) || provide.equals(0L) ? defaultVal : provide;
    }

    /**
     * 如果provide为{@code null}或{@code 0},则返回defaultVal
     *
     * @param provide
     * @param defaultVal
     * @return
     */
    public static Integer getIntDefault(Integer provide, Integer defaultVal) {
        return Objects.isNull(provide) || provide.equals(0) ? defaultVal : provide;
    }

    /**
     * 如果provide为{@code null}或{@code 0.0},则返回defaultVal
     *
     * @param provide
     * @param defaultVal
     * @return
     */
    public static Double getDoubleDefault(Double provide, Double defaultVal) {
        return Objects.isNull(provide) || provide.equals(0.0) ? defaultVal : provide;
    }

    /**
     * 如果provide为{@code null}或{@code 0.0F},则返回defaultVal
     *
     * @param provide
     * @param defaultVal
     * @return
     */
    public static Float getFloatDefault(Float provide, Float defaultVal) {
        return Objects.isNull(provide) || provide.equals(0.0F) ? defaultVal : provide;
    }

    /**
     * predicate ? defaultVal : provide
     *
     * @param provide
     * @param predicate
     * @param defaultVal
     * @param <T>
     * @return
     */
    public static <T> T getRefDefault(T provide, Predicate<T> predicate, T defaultVal) {
        return predicate.test(provide) ? defaultVal : provide;
    }
}
