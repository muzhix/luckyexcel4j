package net.hanbd.luckyexcel4j.jackson;


import com.fasterxml.jackson.core.JsonParser;
import com.fasterxml.jackson.core.TreeNode;
import com.fasterxml.jackson.databind.DeserializationContext;
import com.fasterxml.jackson.databind.JsonDeserializer;

import java.io.IOException;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Set;

/**
 * Json反序列化时提取Map中的key为Set
 *
 * @author hanbd
 */
public class IntKeyMapToKeySetDeserializer extends JsonDeserializer<Set<Integer>> {

    @Override
    public Set<Integer> deserialize(JsonParser p, DeserializationContext ctx) throws IOException {
        HashSet<Integer> keySet = new HashSet<>();
        TreeNode node = p.readValueAsTree();
        if (node.isObject()) {
            Iterator<String> keysItr = node.fieldNames();
            while (keysItr.hasNext()) {
                keySet.add(Integer.valueOf(keysItr.next()));
            }
        }
        return keySet;
    }
}
