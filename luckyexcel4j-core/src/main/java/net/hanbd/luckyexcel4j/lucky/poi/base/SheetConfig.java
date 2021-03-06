package net.hanbd.luckyexcel4j.lucky.poi.base;

import com.fasterxml.jackson.annotation.JsonProperty;
import com.fasterxml.jackson.core.JsonGenerator;
import com.fasterxml.jackson.core.JsonParser;
import com.fasterxml.jackson.core.ObjectCodec;
import com.fasterxml.jackson.core.TreeNode;
import com.fasterxml.jackson.databind.*;
import com.fasterxml.jackson.databind.annotation.JsonDeserialize;
import com.fasterxml.jackson.databind.annotation.JsonSerialize;
import com.fasterxml.jackson.databind.node.ArrayNode;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
import net.hanbd.luckyexcel4j.jackson.IntKeyMapToKeySetDeserializer;
import net.hanbd.luckyexcel4j.jackson.IntSetToZeroValMapSerializer;
import net.hanbd.luckyexcel4j.lucky.poi.enums.BorderRangeType;
import org.apache.commons.compress.utils.Lists;

import java.io.IOException;
import java.util.*;

/**
 * LuckySheet config属性
 *
 * @author hanbd
 * @see <a
 * href="https://mengshukeji.gitee.io/LuckysheetDocs/zh/guide/sheet.html#config-rowlen">rowHeight</a>
 * @see <a
 * href="https://mengshukeji.gitee.io/LuckysheetDocs/zh/guide/sheet.html#config-columnlen">colWidth</a>
 * @see <a
 * href="https://mengshukeji.gitee.io/LuckysheetDocs/zh/guide/sheet.html#config-rowhidden">rowHidden</a>
 * @see <a
 * href="https://mengshukeji.gitee.io/LuckysheetDocs/zh/guide/sheet.html#config-colhidden">colHidden</a>
 */
@Data
@Builder
@AllArgsConstructor
@NoArgsConstructor
public class SheetConfig {
    /**
     * 合并单元格信息.
     * <p>
     * <code>key</code>为 <code>r + '_' + c</code>.
     * r: {@link MergeCell#getStartRow()}, c: {@link MergeCell#getStartCol()}
     *
     * @see MergeCell#mainCellStr
     */
    private Map<String, MergeCell> merge;
    /**
     * 边框信息
     */
    @JsonProperty("borderInfo")
    @JsonSerialize(using = SheetConfigBorderInfoSerializer.class)
    @JsonDeserialize(using = SheetConfigBorderInfoDeserializer.class)
    private List<Border> borderInfos;
    /**
     * 行高信息. key: row index, value: row height
     */
    @JsonProperty("rowlen")
    private Map<Integer, Integer> rowHeight;
    /**
     * 列宽信息. key: column index, value: column width
     */
    @JsonProperty("columnlen")
    private Map<Integer, Integer> columnWidth;
    /**
     * 隐藏行. key: index of hidden row, value: 0,
     *
     * <p>key对应行被隐藏,value始终为0.
     */
    @JsonProperty("rowhidden")
    @JsonSerialize(using = IntSetToZeroValMapSerializer.class)
    @JsonDeserialize(using = IntKeyMapToKeySetDeserializer.class)
    private Set<Integer> hiddenRows;
    /**
     * 隐藏列. key: column index
     *
     * <p>key对应列被隐藏,value始终为0.
     */
    @JsonProperty("colhidden")
    @JsonSerialize(using = IntSetToZeroValMapSerializer.class)
    @JsonDeserialize(using = IntKeyMapToKeySetDeserializer.class)
    private Set<Integer> hiddenColumns;

    /**
     * 工作表保护，可以设置当前整个工作表不允许编辑或者部分区域不可编辑,
     * 如果要申请编辑权限需要输入密码，自定义配置用户可以操作的类型等
     * TODO 待实现
     */
    private Object authority;

    public static class SheetConfigBorderInfoSerializer extends JsonSerializer<List<Border>> {
        @Override
        public void serialize(List<Border> borderInfo, JsonGenerator gen, SerializerProvider serializers) throws IOException {
            gen.writeStartArray();
            for (Border border : borderInfo) {
                if (border instanceof CellBorder) {
                    gen.writeObject((CellBorder) border);
                } else if (border instanceof RangeBorder) {
                    gen.writeObject((RangeBorder) border);
                }
            }
            gen.writeEndArray();
        }
    }

    public static class SheetConfigBorderInfoDeserializer extends JsonDeserializer<List<Border>> {
        @Override
        public List<Border> deserialize(JsonParser p, DeserializationContext ctx) throws IOException {
            ObjectCodec codec = p.getCodec();
            List<Border> borderInfo = null;
            TreeNode node = p.readValueAsTree();
            if (node.isArray()) {
                borderInfo = new ArrayList<>(node.size());
                for (JsonNode jsonNode : (ArrayNode) node) {
                    String rangeType = jsonNode.get("rangeType").textValue();
                    if (Objects.equals(rangeType, BorderRangeType.CELL.getType())) {
                        CellBorder cellBorder = jsonNode.traverse(codec).readValueAs(CellBorder.class);
                        borderInfo.add(cellBorder);
                    } else if (Objects.equals(rangeType, BorderRangeType.RANGE.getType())) {
                        RangeBorder rangeBorder = jsonNode.traverse(codec).readValueAs(RangeBorder.class);
                        borderInfo.add(rangeBorder);
                    }
                }
            }
            if (Objects.isNull(borderInfo)) {
                borderInfo = Lists.newArrayList();
            }
            return borderInfo;
        }
    }

}
