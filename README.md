# UniverJson2Excel

将 Univer Sheets 工作簿快照（JSON）转换为 Excel（.xlsx）的 Java 工具类。

核心类：`com.test.UniverJson2Excel`

## 功能特性

- 支持 Univer 工作簿 JSON 中的 `sheets`、`sheetOrder`、`styles`
- 支持单元格数据：
  - 普通值
  - 字符串 / 数字 / 布尔 / 强制文本类型
  - 公式（`f`），同时写入 Univer 中已有的计算结果
  - 富文本（`p`）
- 支持样式：
  - 字体、字号、粗体、斜体、下划线、删除线、上下标
  - 背景色、字体颜色
  - 上下左右边框
  - 水平 / 垂直对齐
  - 自动换行、文字旋转、缩进
  - 数字格式
- 支持工作表：
  - 名称（自动处理 Excel 非法字符和长度限制）
  - 隐藏状态、Tab 颜色、缩放、网格线
  - 冻结窗格（`freeze`）
  - 行高、列宽、隐藏行列
  - 合并单元格
- 支持 `resources` 插件数据：
  - `SHEET_FILTER_PLUGIN`：自动筛选
  - `SHEET_NOTE_PLUGIN`：批注
  - `SHEET_DRAWING_PLUGIN`：图片
  - `SHEET_CONDITIONAL_FORMATTING_PLUGIN`：条件格式（`highlightCell`、`colorScale`、`dataBar`、`iconSet`）
  - `SHEET_PAGE_LAYOUT_PLUGIN`：页面布局（打印区域、页边距、纸张方向/大小、打印缩放、分页符），兼容旧版 `SHEET_PRINT_AREA_PLUGIN`
- 支持导出 PDF：
  - Univer JSON 直接导出 PDF
  - 已有 Excel 文件导出 PDF
  - PDF 会使用打印区域、页边距、纸张方向/大小、打印缩放和分页符

## 环境要求

- JDK 21（`pom.xml` 中 `maven.compiler.source/target = 21`）
- Maven 3.6+
- Aspose.Cells 本地 jar：
  - 路径：`src/main/resources/libs/aspose-cells-25.10.jar`
  - 已在 `pom.xml` 中以 `system` scope 引用
- Jackson 依赖由 Maven 自动下载

## 快速开始

### 1. 构建

```bash
mvn -q -DskipTests package
```

如果需要在命令行运行示例，先生成 classpath 文件：

```bash
mvn -q -DskipTests compile dependency:build-classpath \
  -Dmdep.outputFile=target/classpath.txt
```

### 2. 使用示例

运行内置示例（`Main` 无参数时默认转换 `univerData.json`）：

```bash
java -cp "target/classes;$(cat target/classpath.txt)" com.test.Main \
  src/main/resources/data/univerData2.json \
  target/univer-output2.xlsx
```

输出 PDF 时，把输出路径写成 `.pdf`，`Main` 会自动调用 PDF 导出方法：

```bash
java -cp "target/classes;$(cat target/classpath.txt)" com.test.Main \
  src/main/resources/data/univerData.json \
  target/univer-output.pdf
```

在代码中使用：

```java
import com.test.UniverJson2Excel;

public class Demo {
    public static void main(String[] args) throws Exception {
        // 方式1：Univer JSON 字符串 -> Excel 文件
        String univerJson = "..."; // Univer 工作簿快照 JSON
        UniverJson2Excel.convertToFile(univerJson, "output.xlsx");

        // 方式2：Univer JSON 文件 -> Excel 文件
        UniverJson2Excel.convertToFile(
            new java.io.File("univerData.json"),
            "output.xlsx"
        );

        // 方式3：输入流 -> 输出流（适合 Web 下载场景）
        try (java.io.InputStream in = new java.io.FileInputStream("univerData.json");
             java.io.OutputStream out = new java.io.FileOutputStream("output.xlsx")) {
            UniverJson2Excel.convertToOutputStream(in, out);
        }

        // 方式4：转为 xlsx byte[]，便于存储或下载
        byte[] xlsxBytes = UniverJson2Excel.toXlsxBytes(univerJson);

        // 方式5：Univer JSON -> PDF 文件（会使用打印区域、纸张方向、页边距、分页符）
        UniverJson2Excel.convertToPdf(univerJson, "output.pdf");

        // 方式6：Univer JSON -> PDF byte[]，适合 Web 下载
        byte[] pdfBytes = UniverJson2Excel.toPdfBytes(univerJson);

        // 方式7：已有 Excel 文件 -> PDF 文件
        UniverJson2Excel.excelToPdf("output.xlsx", "output.pdf");

        // 方式8：构建 Aspose.Cells Workbook 后继续二次加工
        com.aspose.cells.Workbook workbook = UniverJson2Excel.toWorkbook(univerJson);
        // 对 workbook 做额外操作...
        workbook.dispose();
    }
}
```

## API 说明

| 方法 | 说明 |
| --- | --- |
| `convertToFile(String univerJson, String outputPath)` | 将 Univer JSON 字符串转换为 xlsx 文件 |
| `convertToFile(File univerJsonFile, String outputPath)` | 将 Univer JSON 文件转换为 xlsx 文件 |
| `convertToOutputStream(InputStream univerJson, OutputStream excelOut)` | 将 Univer JSON 输入流写入 Excel 输出流 |
| `toXlsxBytes(String univerJson)` | 将 Univer JSON 字符串转换为 xlsx 字节数组 |
| `convertToPdf(String univerJson, String outputPath)` | 将 Univer JSON 字符串转换为 PDF 文件 |
| `convertToPdf(File univerJsonFile, String outputPath)` | 将 Univer JSON 文件转换为 PDF 文件 |
| `convertToPdf(InputStream univerJson, OutputStream pdfOut)` | 将 Univer JSON 输入流写入 PDF 输出流 |
| `toPdfBytes(String univerJson)` | 将 Univer JSON 字符串转换为 PDF 字节数组 |
| `excelToPdf(String excelPath, String pdfPath)` | 将已有 Excel 文件导出为 PDF 文件 |
| `excelToPdf(InputStream excelIn, OutputStream pdfOut)` | 将已有 Excel 输入流导出为 PDF 输出流 |
| `excelToPdfBytes(InputStream excelIn)` | 将已有 Excel 输入流导出为 PDF 字节数组 |
| `toWorkbook(String univerJson)` | 解析 JSON 并构建 Aspose.Cells Workbook |
| `toWorkbook(InputStream univerJson)` | 从输入流解析 JSON 并构建 Aspose.Cells Workbook |

> 注意：`Workbook` 使用完毕后应调用 `workbook.dispose()` 释放资源。工具类的文件/流转换方法内部已自动处理。

## 支持范围对照

| Univer 数据 | 生成的 Excel 内容 |
| --- | --- |
| `sheetOrder` / `sheets` | 工作表名称与顺序 |
| `styles` + 单元格 `s` | 字体、颜色、边框、对齐等样式 |
| `cellData` | 单元格值、公式、富文本 |
| `mergeData` | 合并单元格 |
| `rowData` / `columnData` | 行高、列宽、隐藏行列 |
| `freeze` | 冻结窗格 |
| `hidden` / `tabColor` / `zoomRatio` / `showGridlines` | 工作表显示属性 |
| `SHEET_FILTER_PLUGIN` | 自动筛选 |
| `SHEET_NOTE_PLUGIN` | 批注 |
| `SHEET_DRAWING_PLUGIN` | 图片 |
| `SHEET_CONDITIONAL_FORMATTING_PLUGIN` | 条件格式（`highlightCell`、`colorScale`、`dataBar`、`iconSet`） |
| `SHEET_PAGE_LAYOUT_PLUGIN` | 页面布局：打印区域、页边距、纸张方向/大小、打印缩放（`scaleMode`）、分页符 |

## 条件格式说明

- `highlightCell`：
  - 数字/文本比较映射为 Excel `CellIs` 条件格式
  - 支持 `equal`、`notEqual`、`greaterThan`、`greaterThanOrEqual`、`lessThan`、`lessThanOrEqual`、`between`、`notBetween`
  - 自动解析 Univer 的 `&quot;...&quot;` 文本字面量
  - 支持子类型：`rank`（前 N 项/后 N 项）、`uniqueValues`（唯一值）、`duplicateValues`（重复值）、`formula`（公式）
- 条件格式样式兼容性：
  - `highlightCell`（如 `rank`、`greaterThan`、`lessThan` 等）的背景色会同时写入 DXF 的 `fgColor` 和 `bgColor`，避免部分 Excel/WPS 版本不显示单元格底色
- `colorScale`：
  - 支持双色和三色色阶
  - 支持 `min` / `max` / `percent` / `percentile` / `formula` 取值类型
  - 兼容颜色字符串 `"#FF7128"` 和颜色对象 `{"rgb":"#FF7128"}`
- `dataBar`：
  - 使用 `config.min` / `config.max` 设置数据条最小/最大值
  - 支持 `isGradient`（渐变/实心）、`isShowValue`（是否显示数值）
  - `positiveColor` 映射为数据条颜色，`nativeColor` 映射为负值数据条颜色
- `iconSet`：
  - 支持 `3Arrows`、`4Arrows`、`5Arrows`、`3TrafficLights1`、`3TrafficLights2`、`3Symbols`、`3Flags`、`3Stars` 等常见图标集
  - 按 Univer 的 `config` 顺序逆序写入 Excel 阈值（Excel 阈值从低到高排列）
  - 支持 `num` / `percent` / `percentile` / `formula` 等取值类型
  - 部分图标集（如 `3Triangles`、`3Stars` 等）会以 x14 扩展条件格式写入，Excel 2010+ 可正常显示

## 页面布局说明

`SHEET_PAGE_LAYOUT_PLUGIN` 的 `data` 按工作表 ID 保存页面布局，支持以下字段：

- `printArea`：`IRange` 数组，转换为 Excel 打印区域；多个区域会用英文逗号连接
- `margins`：页边距，单位为英寸
  - `top` / `bottom` / `left` / `right` / `header` / `footer`
- `pageSetup`：
  - `orientation`：`landscape` / `portrait`
  - `paperSize`：Excel 纸张大小编号，如 `9` 表示 A4
  - `scaleMode`：打印缩放方式
  - `scale`：普通缩放百分比，仅 `scaleMode = none` 时使用，范围 10-400
- `rowBreaks`：水平分页符行索引数组
- `columnBreaks`：垂直分页符列索引数组

`scaleMode` 映射关系：

| Univer scaleMode | 含义 | Excel 设置 |
| --- | --- | --- |
| `none` | 不缩放，使用普通缩放比例 | 关闭“调整为”，`scale` 默认 100 |
| `fitSheet` | 整个工作表缩放到一页 | 宽 1 页、高 1 页 |
| `fitColumns` | 所有列缩放到一页宽，行可跨页 | 宽 1 页、高不限 |
| `fitRows` | 所有行缩放到一页高，列可跨页 | 宽不限、高 1 页 |

示例：

```json
{
  "sheet-1": {
    "printArea": [
      { "startRow": 0, "startColumn": 0, "endRow": 35, "endColumn": 8 }
    ],
    "margins": {
      "top": 0.75,
      "bottom": 0.75,
      "left": 0.7,
      "right": 0.7,
      "header": 0.3,
      "footer": 0.3
    },
    "pageSetup": {
      "orientation": "landscape",
      "paperSize": 9,
      "scaleMode": "fitRows"
    },
    "rowBreaks": [20]
  }
}
```

## 当前未支持（可继续扩展）

- `SHEET_DATA_VALIDATION_PLUGIN`：数据验证
- `SHEET_TABLE_PLUGIN`：表格
- `SHEET_DEFINED_NAME_PLUGIN`：定义名称
- `SHEET_WORKSHEET_PROTECTION_PLUGIN`：工作表保护
- `SHEET_UNIVER_THREAD_COMMENT_PLUGIN`：论坛批注
- 条件格式中 `highlightCell` 的更多子类型：文本/日期/平均值
- 超链接、OLE 附件等

## 示例数据

- `src/main/resources/data/univerData.json`：基础示例（含公式、富文本、批注、图片、筛选、页面布局，以及色阶/数据条/图标集/前 N 项/大于条件格式）
- `src/main/resources/data/univerData2.json`：扩展示例（含页面布局 `scaleMode = fitRows`、打印区域、`highlightCell` 条件格式）

## 最近更新

- 条件格式新增 `dataBar`、`iconSet` 支持
- `highlightCell` 新增子类型支持：`rank`（前 N 项/后 N 项）、`uniqueValues`、`duplicateValues`、`formula`
- 修复 `iconSet` 阈值类型和顺序还原（如 `percent` 类型的 `3Triangles` 等）
- 修复 `Top10` 条件格式被错误还原为 `CellIs equal` 的问题
- 修复条件格式背景色兼容性：DXF 同时写入 `fgColor` 与 `bgColor`，避免部分 Excel/WPS 不显示单元格底色
- 适配 `SHEET_PAGE_LAYOUT_PLUGIN` 页面布局：打印区域、页边距、纸张方向/大小、打印缩放（`scaleMode`）、行/列分页符
- 新增 PDF 导出能力：`convertToPdf`、`toPdfBytes`、`excelToPdf`、`excelToPdfBytes`

## 注意事项

- Aspose.Cells 使用本地 jar，未发布到 Maven Central；迁移环境时需要保留
  `src/main/resources/libs/aspose-cells-25.10.jar` 文件并保持 `pom.xml` 中的路径一致。
- 工作表名称会做 Excel 合法性处理：非法字符替换为 `_`，超过 31 个字符截断。
