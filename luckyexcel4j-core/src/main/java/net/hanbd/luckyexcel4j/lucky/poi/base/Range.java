package net.hanbd.luckyexcel4j.lucky.poi.base;

import com.fasterxml.jackson.core.JsonGenerator;
import com.fasterxml.jackson.core.JsonParser;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.DeserializationContext;
import com.fasterxml.jackson.databind.JsonDeserializer;
import com.fasterxml.jackson.databind.JsonSerializer;
import com.fasterxml.jackson.databind.SerializerProvider;
import com.fasterxml.jackson.databind.annotation.JsonDeserialize;
import com.fasterxml.jackson.databind.annotation.JsonSerialize;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.io.IOException;
import java.util.List;

/**
 * 选中区域
 *
 * @author hanbd
 */
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class Range {

    public Range(Integer rowIndex, Integer columnIndex) {
        this.row = new StartEndPair(rowIndex, rowIndex);
        this.column = new StartEndPair(columnIndex, columnIndex);
    }

    /**
     * 行起止index范围
     */
    @JsonSerialize(using = StartEndPairSerializer.class)
    @JsonDeserialize(using = StartEndPairDeserializer.class)
    private StartEndPair row;
    /**
     * 列起止index范围
     */
    @JsonSerialize(using = StartEndPairSerializer.class)
    @JsonDeserialize(using = StartEndPairDeserializer.class)
    private StartEndPair column;

    /**
     * 起止index对
     */
    @Data
    @Builder
    @AllArgsConstructor
    static class StartEndPair {
        /**
         * start index
         */
        private Integer start;
        /**
         * end index
         */
        private Integer end;
    }

    static class StartEndPairSerializer extends JsonSerializer<StartEndPair> {
        @Override
        public void serialize(StartEndPair value, JsonGenerator gen, SerializerProvider serializers) throws IOException {
            gen.writeStartArray();
            gen.writeNumber(value.getStart());
            gen.writeNumber(value.getEnd());
            gen.writeEndArray();
        }
    }

    static class StartEndPairDeserializer extends JsonDeserializer<StartEndPair> {
        @Override
        public StartEndPair deserialize(JsonParser p, DeserializationContext ctx) throws IOException {
            List<Integer> values = p.readValueAs(new TypeReference<List<Integer>>() {
            });
            StartEndPair.StartEndPairBuilder builder = StartEndPair.builder();
            if (values.size() >= 2) {
                builder.start(values.get(0))
                        .end(values.get(1));
            }
            return builder.build();
        }
    }
}
