package net.hanbd.luckyexcel4j.jackson;

import com.fasterxml.jackson.core.JsonGenerator;
import com.fasterxml.jackson.databind.JsonSerializer;
import com.fasterxml.jackson.databind.SerializerProvider;

import java.io.IOException;
import java.util.Set;

/**
 * @author hanbd
 */
public class IntSetToZeroValMapSerializer extends JsonSerializer<Set<Integer>> {
    @Override
    public void serialize(Set<Integer> set, JsonGenerator gen, SerializerProvider serializers) throws IOException {
        gen.writeStartObject();
        for (Integer integer : set) {
            gen.writeFieldName(String.valueOf(integer));
            gen.writeNumber(0);
        }
        gen.writeEndObject();
    }
}
