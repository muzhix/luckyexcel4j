package net.hanbd.luckyexcel4j.lucky.poi.enums;

import lombok.AllArgsConstructor;
import lombok.Getter;

/**
 * 边框线位置类型
 *
 * @see <a
 *     href="https://mengshukeji.gitee.io/LuckysheetDocs/zh/guide/sheet.html#config-borderinfo">config-borderInfo</a>
 * @author hanbd
 */
@AllArgsConstructor
public enum BorderType {
  /** 左框线 */
  LEFT("border-left"),
  /** 右框线 */
  RIGHT("border-right"),
  /** 上框线 */
  TOP("border-top"),
  /** 下框线 */
  BOTTOM("border-bottom"),
  /** 所有框线 */
  ALL("border-all"),
  /** 外侧框线 */
  OUTSIDE("border-outside"),
  /** 内部框线 */
  INSIDE("border-inside"),
  /** 内部水平框线 */
  HORIZONTAL("border-horizontal"),
  /** 内部纵向框线 */
  VERTICAL("border-vertical"),
  /** 无框线 */
  NONE("border-none");

  @Getter private final String type;
}