package net.hanbd.luckyexcel4j.utils;

import com.fasterxml.jackson.core.JsonGenerator;
import com.fasterxml.jackson.core.JsonParser;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.*;

import java.io.File;
import java.io.IOException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Map;

/**
 * POJO <--> JSON
 *
 * @author hanbd
 * @date 2020/09/03
 */
public class JsonUtil {
    private JsonUtil() {
        throw new AssertionError();
    }

    private static final ObjectMapper MAPPER = new ObjectMapper();

    /**
     * {@link LocalDateTime} 序列化
     */
    public static class LocalDateTimeSerializer extends JsonSerializer<LocalDateTime> {
        @Override
        public void serialize(LocalDateTime localDateTime, JsonGenerator jsonGenerator, SerializerProvider serializerProvider) throws IOException {
            String localDateTimeStr = localDateTime.format(DateTimeFormatter.ISO_DATE_TIME);
            jsonGenerator.writeString(localDateTimeStr);
        }
    }

    /**
     * {@link LocalDateTime} 反序列化
     */
    public static class LocalDateTimeDeserializer extends JsonDeserializer<LocalDateTime> {
        @Override
        public LocalDateTime deserialize(JsonParser jsonParser, DeserializationContext deserializationContext) throws IOException, JsonProcessingException {
            String str = jsonParser.getText().trim();
            return LocalDateTime.parse(str, DateTimeFormatter.ISO_DATE_TIME);
        }
    }

    /**
     * {@link LocalDate} 序列化
     */
    public static class LocalDateSerializer extends JsonSerializer<LocalDate> {
        @Override
        public void serialize(LocalDate localDate, JsonGenerator jsonGenerator, SerializerProvider serializerProvider) throws IOException {
            String localDateStr = localDate.format(DateTimeFormatter.ISO_DATE);
            jsonGenerator.writeString(localDateStr);
        }
    }

    /**
     * {@link LocalDate} 反序列化
     */
    public static class LocalDateDeserializer extends JsonDeserializer<LocalDate> {

        @Override
        public LocalDate deserialize(JsonParser jsonParser, DeserializationContext deserializationContext) throws IOException {
            String str = jsonParser.getText().trim();
            return LocalDate.parse(str, DateTimeFormatter.ISO_DATE);
        }
    }

    /**
     * POJO -> JSON
     *
     * @param obj pojo
     * @return 转换成功返回json字符串, 转换失败返回{@code ""}
     */
    public static String stringify(Object obj) {
        String json = "";
        try {
            json = MAPPER.writeValueAsString(obj);
        } catch (JsonProcessingException e) {
            e.printStackTrace();
        }
        return json;
    }

    /**
     * JSON str -> POJO
     *
     * @param json  json字符串
     * @param clazz 转换目标类型
     * @param <T>
     * @return 转换成功返回对应对象, 转换失败返回{@code null}
     */
    public static <T> T parse(String json, Class<T> clazz) {
        T object = null;
        try {
            object = MAPPER.readValue(json, clazz);
        } catch (JsonProcessingException e) {
            e.printStackTrace();
        }
        return object;
    }

    /**
     * JSON file -> POJO
     *
     * @param jsonFile json文件
     * @param clazz    转换目标类型
     * @param <T>
     * @return 转换成功返回对应对象, 转换失败返回{@code null}
     */
    public static <T> T parse(File jsonFile, Class<T> clazz) {
        T object = null;
        try {
            object = MAPPER.readValue(jsonFile, clazz);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return object;
    }

    /**
     * JSON str -> MAP
     *
     * @param json json字符串
     * @return 转换成功返回对应MAP, 转换失败返回{@code null}
     */
    public static Map<String, Object> parse2Map(String json) {
        Map<String, Object> map = null;
        try {
            map = MAPPER.readValue(json, new TypeReference<Map<String, Object>>() {
            });
        } catch (JsonProcessingException e) {
            e.printStackTrace();
        }
        return map;
    }

}
