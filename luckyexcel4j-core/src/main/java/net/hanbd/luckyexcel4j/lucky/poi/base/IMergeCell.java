package net.hanbd.luckyexcel4j.lucky.poi.base;

import com.fasterxml.jackson.core.JsonGenerator;
import com.fasterxml.jackson.core.JsonParser;
import com.fasterxml.jackson.core.TreeNode;
import com.fasterxml.jackson.databind.DeserializationContext;
import com.fasterxml.jackson.databind.JsonDeserializer;
import com.fasterxml.jackson.databind.JsonSerializer;
import com.fasterxml.jackson.databind.SerializerProvider;

import java.io.IOException;
import java.util.Objects;

/**
 * 合并单元格标记接口。使用该接口主要是为了处理住主单元格与合并单元格内其他单元格的赋值问题
 *
 * @author hanbd
 * @see
 * <a href="https://mengshukeji.gitee.io/LuckysheetDocs/zh/guide/cell.html#%E5%90%88%E5%B9%B6%E5%8D%95%E5%85%83%E6%A0%BC">
 * 合并单元格处理</a>
 */
@Deprecated
public interface IMergeCell {
    class IMergeCellSerializer extends JsonSerializer<IMergeCell> {
        @Override
        public void serialize(IMergeCell value, JsonGenerator gen, SerializerProvider serializers) throws IOException {
            if (value instanceof MergeCell) {
                gen.writeObject((MergeCell) value);
            } else if (value instanceof MergeCell.MainCell) {
                gen.writeObject((MergeCell.MainCell) value);
            }
        }
    }

    class IMergeCellDeSerializer extends JsonDeserializer<IMergeCell> {

        @Override
        public IMergeCell deserialize(JsonParser p, DeserializationContext ctx) throws IOException {
            IMergeCell mergeCell = null;
            TreeNode node = p.readValueAsTree();
            if (node.isObject()) {
                if (Objects.nonNull(node.get("rs")) && Objects.nonNull(node.get("cs"))) {
                    mergeCell = p.readValueAs(MergeCell.class);
                } else {
                    mergeCell = p.readValueAs(MergeCell.MainCell.class);
                }
            }
            return mergeCell;
        }
    }
}
