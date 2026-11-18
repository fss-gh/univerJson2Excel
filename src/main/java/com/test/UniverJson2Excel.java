package com.test;

import com.aspose.cells.BackgroundType;
import com.aspose.cells.BorderType;
import com.aspose.cells.Cell;
import com.aspose.cells.CellBorderType;
import com.aspose.cells.Cells;
import com.aspose.cells.Color;
import com.aspose.cells.DataBar;
import com.aspose.cells.DataBarFillType;
import com.aspose.cells.DataBarNegativeColorType;
import com.aspose.cells.ColorScale;
import com.aspose.cells.Comment;
import com.aspose.cells.ConditionalFormattingValue;
import com.aspose.cells.Font;
import com.aspose.cells.FontSetting;
import com.aspose.cells.FontUnderlineType;
import com.aspose.cells.FormatCondition;
import com.aspose.cells.FormatConditionCollection;
import com.aspose.cells.FormatConditionType;
import com.aspose.cells.FormatConditionValueType;
import com.aspose.cells.IconSet;
import com.aspose.cells.IconSetType;
import com.aspose.cells.CellArea;
import com.aspose.cells.OperatorType;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Style;
import com.aspose.cells.TextAlignmentType;
import com.aspose.cells.VisibilityType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Base64;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * 将 Univer Sheets 工作簿快照(JSON)转换为 Excel 的工具类。
 *
 * <p>支持的数据范围：</p>
 * <ul>
 *   <li>工作簿 sheets / sheetOrder / styles</li>
 *   <li>单元格：值、公式、富文本、样式</li>
 *   <li>样式：字体、字号、粗体、斜体、下划线、删除线、上下标、
 *       背景色、字体色、边框、对齐、自动换行、数字格式、文字旋转</li>
 *   <li>合并单元格 mergeData</li>
 *   <li>行高 rowData、列宽 columnData、隐藏行列</li>
 *   <li>冻结 freeze、工作表隐藏、标题颜色、缩放、网格线</li>
 *   <li>resources 中的 SHEET_FILTER_PLUGIN(自动筛选)、SHEET_NOTE_PLUGIN(批注)、
 *       SHEET_DRAWING_PLUGIN(图片)、SHEET_CONDITIONAL_FORMATTING_PLUGIN(条件格式)、
 *       SHEET_PRINT_AREA_PLUGIN(打印区域)</li>
 * </ul>
 */
public final class UniverJson2Excel {

    private static final ObjectMapper MAPPER = new ObjectMapper();

    private UniverJson2Excel() {
    }

    /**
     * 将 Univer JSON 字符串转换为 Excel 文件。
     *
     * @param univerJson Univer 工作簿 JSON
     * @param outputPath 输出 xlsx 路径
     */
    public static void convertToFile(String univerJson, String outputPath) throws Exception {
        Workbook workbook = toWorkbook(univerJson);
        try {
            workbook.save(outputPath, SaveFormat.XLSX);
        } finally {
            workbook.dispose();
        }
    }

    /**
     * 将 Univer JSON 文件转换为 Excel 文件。
     */
    public static void convertToFile(File univerJsonFile, String outputPath) throws Exception {
        try (InputStream in = new FileInputStream(univerJsonFile);
             OutputStream out = new java.io.FileOutputStream(outputPath)) {
            convertToOutputStream(in, out);
        }
    }

    /**
     * 将 Univer JSON 输入流写入 Excel 输出流。
     */
    public static void convertToOutputStream(InputStream univerJson, OutputStream excelOut) throws Exception {
        Workbook workbook = toWorkbook(univerJson);
        try {
            workbook.save(excelOut, SaveFormat.XLSX);
        } finally {
            workbook.dispose();
        }
    }

    /**
     * 将 Univer JSON 字符串转换为 xlsx 字节数组。
     */
    public static byte[] toXlsxBytes(String univerJson) throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        convertToOutputStream(new java.io.ByteArrayInputStream(univerJson.getBytes(java.nio.charset.StandardCharsets.UTF_8)), out);
        return out.toByteArray();
    }

    /**
     * 将 Univer 工作簿 JSON 解析并构建为 Aspose.Cells Workbook。
     */
    public static Workbook toWorkbook(String univerJson) throws Exception {
        JsonNode root = MAPPER.readTree(univerJson);
        return toWorkbook(root);
    }

    /**
     * 将 Univer 工作簿 JSON 输入流解析并构建为 Aspose.Cells Workbook。
     */
    public static Workbook toWorkbook(InputStream univerJson) throws Exception {
        JsonNode root = MAPPER.readTree(univerJson);
        return toWorkbook(root);
    }

    private static Workbook toWorkbook(JsonNode root) throws IOException {
        Workbook workbook = new Workbook();
        fillWorkbook(workbook, root);
        return workbook;
    }

    // ------------------------------------------------------------------
    // Workbook / Worksheet
    // ------------------------------------------------------------------

    private static void fillWorkbook(Workbook workbook, JsonNode root) {
        JsonNode sheetsNode = root.path("sheets");
        JsonNode stylesNode = root.path("styles");

        List<String> sheetOrder = new ArrayList<>();
        JsonNode orderNode = root.path("sheetOrder");
        if (orderNode.isArray()) {
            for (JsonNode id : orderNode) {
                sheetOrder.add(id.asText());
            }
        }
        if (sheetOrder.isEmpty() && sheetsNode.isObject()) {
            sheetsNode.fieldNames().forEachRemaining(sheetOrder::add);
        }

        WorksheetCollection worksheets = workbook.getWorksheets();
        worksheets.clear();

        Map<String, Worksheet> sheetMap = new LinkedHashMap<>();
        int defaultIndex = 0;
        for (String sheetId : sheetOrder) {
            JsonNode sheetNode = sheetsNode.path(sheetId);
            if (!sheetNode.isObject()) {
                continue;
            }
            String sheetName = sheetNode.path("name").asText("");
            if (sheetName.isEmpty()) {
                sheetName = "Sheet" + (defaultIndex + 1);
            }
            Worksheet worksheet = addWorksheet(worksheets, sheetName, defaultIndex);
            sheetMap.put(sheetId, worksheet);
            fillWorksheet(workbook, worksheet, sheetNode, stylesNode);
            defaultIndex++;
        }

        if (defaultIndex == 0) {
            worksheets.add("Sheet1");
        }

        applyResources(workbook, root.path("resources"), sheetMap);
    }

    private static Worksheet addWorksheet(WorksheetCollection worksheets, String baseName, int index) {
        String name = sanitizeSheetName(baseName);
        if (name.isEmpty()) {
            name = "Sheet" + (index + 1);
        }
        try {
            return worksheets.add(name);
        } catch (Exception e) {
            String uniqueName = name + "_" + (index + 1);
            try {
                return worksheets.add(uniqueName);
            } catch (Exception ex) {
                return worksheets.add("Sheet" + (index + 1));
            }
        }
    }

    private static String sanitizeSheetName(String name) {
        if (name == null) {
            return "";
        }
        String s = name.trim();
        // Excel 工作表名称不允许 : \ / ? * [ ]
        s = s.replaceAll("[:\\\\/?*\\[\\]]", "_");
        if (s.length() > 31) {
            s = s.substring(0, 31);
        }
        return s;
    }

    private static void fillWorksheet(Workbook workbook, Worksheet worksheet, JsonNode sheetNode, JsonNode globalStyles) {
        // 工作表属性
        if (sheetNode.has("hidden")) {
            int hidden = sheetNode.path("hidden").asInt(0);
            worksheet.setVisibilityType(hidden == 2 ? VisibilityType.VERY_HIDDEN
                    : hidden == 1 ? VisibilityType.HIDDEN : VisibilityType.VISIBLE);
        }
        if (sheetNode.has("tabColor")) {
            String tabColor = sheetNode.path("tabColor").asText("");
            if (!tabColor.isEmpty()) {
                Color color = parseColor(tabColor);
                if (color != null) {
                    worksheet.setTabColor(color);
                }
            }
        }
        if (sheetNode.has("zoomRatio")) {
            worksheet.setZoom((int) Math.round(sheetNode.path("zoomRatio").asDouble(1.0) * 100));
        }
        if (sheetNode.has("showGridlines")) {
            worksheet.setGridlinesVisible(sheetNode.path("showGridlines").asInt(1) != 0);
        }

        Cells cells = worksheet.getCells();

        // 默认行高 / 列宽(Univer 单位是 px)
        if (sheetNode.has("defaultRowHeight")) {
            cells.setStandardHeightPixels(sheetNode.path("defaultRowHeight").asInt());
        }
        if (sheetNode.has("defaultColumnWidth")) {
            cells.setStandardWidthPixels(sheetNode.path("defaultColumnWidth").asInt());
        }

        // 冻结窗格
        JsonNode freeze = sheetNode.path("freeze");
        if (freeze.isObject()) {
            int xSplit = freeze.path("xSplit").asInt(0);
            int ySplit = freeze.path("ySplit").asInt(0);
            if (xSplit > 0 || ySplit > 0) {
                // Aspose: freezePanes(row, column, freezedRows, freezedColumns)
                worksheet.freezePanes(ySplit, xSplit, ySplit, xSplit);
            }
        }

        // 单元格
        JsonNode cellData = sheetNode.path("cellData");
        if (cellData.isObject()) {
            Iterator<Map.Entry<String, JsonNode>> rowIt = cellData.fields();
            while (rowIt.hasNext()) {
                Map.Entry<String, JsonNode> rowEntry = rowIt.next();
                int row = parseInt(rowEntry.getKey(), -1);
                if (row < 0) {
                    continue;
                }
                JsonNode rowCells = rowEntry.getValue();
                if (!rowCells.isObject()) {
                    continue;
                }
                Iterator<Map.Entry<String, JsonNode>> colIt = rowCells.fields();
                while (colIt.hasNext()) {
                    Map.Entry<String, JsonNode> colEntry = colIt.next();
                    int col = parseInt(colEntry.getKey(), -1);
                    if (col < 0) {
                        continue;
                    }
                    JsonNode cellNode = colEntry.getValue();
                    if (!cellNode.isObject()) {
                        continue;
                    }
                    Cell cell = cells.get(row, col);
                    fillCell(workbook, worksheet, cell, cellNode, globalStyles);
                }
            }
        }

        // 行高 / 隐藏行
        JsonNode rowData = sheetNode.path("rowData");
        if (rowData.isObject()) {
            Iterator<Map.Entry<String, JsonNode>> it = rowData.fields();
            while (it.hasNext()) {
                Map.Entry<String, JsonNode> entry = it.next();
                int row = parseInt(entry.getKey(), -1);
                if (row < 0) {
                    continue;
                }
                JsonNode rowNode = entry.getValue();
                if (rowNode.has("h") && rowNode.get("h").isNumber()) {
                    cells.setRowHeightPixel(row, rowNode.get("h").asInt());
                }
                if (rowNode.path("hd").asInt(0) != 0) {
                    cells.getRows().get(row).setHidden(true);
                }
            }
        }

        // 列宽 / 隐藏列
        JsonNode columnData = sheetNode.path("columnData");
        if (columnData.isObject()) {
            Iterator<Map.Entry<String, JsonNode>> it = columnData.fields();
            while (it.hasNext()) {
                Map.Entry<String, JsonNode> entry = it.next();
                int col = parseInt(entry.getKey(), -1);
                if (col < 0) {
                    continue;
                }
                JsonNode colNode = entry.getValue();
                if (colNode.has("w") && colNode.get("w").isNumber()) {
                    cells.setColumnWidthPixel(col, colNode.get("w").asInt());
                }
                if (colNode.path("hd").asInt(0) != 0) {
                    cells.getColumns().get(col).setHidden(true);
                }
            }
        }

        // 合并单元格(放在数据之后,保留左上角的值)
        JsonNode mergeData = sheetNode.path("mergeData");
        if (mergeData.isArray()) {
            for (JsonNode merge : mergeData) {
                if (!merge.isObject()) {
                    continue;
                }
                int startRow = merge.path("startRow").asInt(-1);
                int endRow = merge.path("endRow").asInt(-1);
                int startColumn = merge.path("startColumn").asInt(-1);
                int endColumn = merge.path("endColumn").asInt(-1);
                if (startRow < 0 || endRow < startRow || startColumn < 0 || endColumn < startColumn) {
                    continue;
                }
                cells.merge(startRow, startColumn, endRow - startRow + 1, endColumn - startColumn + 1);
            }
        }
    }

    // ------------------------------------------------------------------
    // Cell
    // ------------------------------------------------------------------

    private static void fillCell(Workbook workbook, Worksheet worksheet, Cell cell, JsonNode cellNode, JsonNode globalStyles) {
        // 样式
        JsonNode styleNode = resolveStyle(cellNode.path("s"), globalStyles);
        if (styleNode != null && styleNode.isObject()) {
            Style style = cell.getStyle();
            applyStyle(style, styleNode);
            cell.setStyle(style);
        }

        JsonNode valueNode = cellNode.get("v");
        JsonNode formulaNode = cellNode.get("f");
        JsonNode richTextNode = cellNode.get("p");
        int type = cellNode.path("t").asInt(0);

        // 公式优先
        if (formulaNode != null && formulaNode.isTextual() && !formulaNode.asText().isEmpty()) {
            String formula = formulaNode.asText();
            if (!formula.startsWith("=")) {
                formula = "=" + formula;
            }
            if (isValidValue(valueNode)) {
                cell.setFormula(formula, toValueObject(valueNode, type));
            } else {
                cell.setFormula(formula);
            }
            return;
        }

        // 富文本(p 存在时优先展示 p)
        if (richTextNode != null && richTextNode.isObject()) {
            setRichText(workbook, worksheet, cell, richTextNode);
            return;
        }

        // 普通值
        if (isValidValue(valueNode)) {
            setCellValue(cell, valueNode, type);
        }
    }

    /**
     * 解析样式：cellData.s 可能是样式 id,也可能是内联样式对象。
     */
    private static JsonNode resolveStyle(JsonNode styleRef, JsonNode globalStyles) {
        if (styleRef == null || styleRef.isNull()) {
            return null;
        }
        if (styleRef.isObject()) {
            return styleRef;
        }
        if (styleRef.isValueNode() && globalStyles.isObject()) {
            return globalStyles.get(styleRef.asText());
        }
        return null;
    }

    private static void setCellValue(Cell cell, JsonNode valueNode, int type) {
        switch (type) {
            case 2: // number
                if (valueNode.isNumber()) {
                    cell.putValue(valueNode.doubleValue());
                } else if (valueNode.isTextual()) {
                    String text = valueNode.asText();
                    try {
                        cell.putValue(Double.parseDouble(text));
                    } catch (NumberFormatException e) {
                        cell.putValue(text);
                    }
                } else {
                    cell.putValue(toValueObject(valueNode, type));
                }
                break;
            case 3: // boolean,Univer 用 0/1 表示
                cell.putValue(!isZero(valueNode));
                break;
            case 4: // force text
                cell.putValue(valueNode.asText());
                break;
            default: // string 或自动识别
                if (valueNode.isNumber()) {
                    cell.putValue(valueNode.doubleValue());
                } else {
                    cell.putValue(valueNode.asText());
                }
                break;
        }
    }

    private static Object toValueObject(JsonNode valueNode, int type) {
        if (valueNode == null || valueNode.isNull()) {
            return null;
        }
        if (type == 3) {
            return !isZero(valueNode);
        }
        if (valueNode.isNumber()) {
            return valueNode.doubleValue();
        }
        return valueNode.asText();
    }

    private static boolean isValidValue(JsonNode valueNode) {
        return valueNode != null && !valueNode.isNull() && !valueNode.isMissingNode();
    }

    private static boolean isZero(JsonNode node) {
        if (node.isBoolean()) {
            return !node.asBoolean();
        }
        return node.asDouble(0.0) == 0.0;
    }

    // ------------------------------------------------------------------
    // Rich text
    // ------------------------------------------------------------------

    private static void setRichText(Workbook workbook, Worksheet worksheet, Cell cell, JsonNode richTextNode) {
        JsonNode body = richTextNode.path("body");
        String dataStream = body.path("dataStream").asText("");
        String text = cleanRichTextDataStream(dataStream);
        if (text.isEmpty()) {
            return;
        }

        cell.putValue(text);

        JsonNode textRuns = body.path("textRuns");
        if (!textRuns.isArray()) {
            return;
        }

        List<FontSetting> fontSettings = new ArrayList<>();
        for (JsonNode run : textRuns) {
            int start = run.path("st").asInt(-1);
            int end = run.path("ed").asInt(-1);
            if (start < 0 || end <= start || start >= text.length()) {
                continue;
            }
            int length = Math.min(end - start, text.length() - start);
            if (length <= 0) {
                continue;
            }
            FontSetting fontSetting = new FontSetting(start, length, workbook.getWorksheets());
            JsonNode textStyle = run.get("ts");
            if (textStyle != null && textStyle.isObject()) {
                applyFontStyle(fontSetting.getFont(), textStyle);
            }
            fontSettings.add(fontSetting);
        }

        if (!fontSettings.isEmpty()) {
            cell.setCharacters(fontSettings.toArray(new FontSetting[0]));
        }
    }

    /**
     * Univer 富文本 dataStream 中,\r 是段落标记,\n 是分节标记。
     * 转为 Excel 单元格文本:段落标记变为换行,并去掉末尾的分节/段落标记。
     */
    private static String cleanRichTextDataStream(String dataStream) {
        if (dataStream == null || dataStream.isEmpty()) {
            return "";
        }
        String text = dataStream.replace('\r', '\n');
        int end = text.length();
        while (end > 0 && text.charAt(end - 1) == '\n') {
            end--;
        }
        return text.substring(0, end);
    }

    // ------------------------------------------------------------------
    // Style
    // ------------------------------------------------------------------

    private static void applyStyle(Style style, JsonNode styleNode) {
        applyFontStyle(style.getFont(), styleNode);

        // 背景色
        JsonNode bg = styleNode.get("bg");
        if (bg != null && bg.isObject()) {
            Color color = parseColor(bg.path("rgb").asText(""));
            if (color != null) {
                style.setPattern(BackgroundType.SOLID);
                style.setForegroundColor(color);
            }
        }

        // 水平对齐
        if (styleNode.has("ht")) {
            style.setHorizontalAlignment(mapHorizontalAlignment(styleNode.get("ht").asInt(0)));
        }

        // 垂直对齐
        if (styleNode.has("vt")) {
            style.setVerticalAlignment(mapVerticalAlignment(styleNode.get("vt").asInt(0)));
        }

        // 自动换行:Univer 1 溢出,2 截断,3 换行
        if (styleNode.has("tb")) {
            style.setTextWrapped(styleNode.get("tb").asInt(0) == 3);
        }

        // 文字旋转
        JsonNode tr = styleNode.get("tr");
        if (tr != null && tr.isObject()) {
            if (tr.path("v").asInt(0) == 1) {
                style.setRotationAngle(255); // Excel 竖排文字
            } else if (tr.has("a")) {
                style.setRotationAngle(tr.get("a").asInt(0));
            }
        }

        // 缩小字体填充
        if (styleNode.has("stf")) {
            style.setShrinkToFit(styleNode.get("stf").asInt(0) != 0);
        }

        // 内边距:Excel 没有完全等效的 API,用左缩进近似
        JsonNode pd = styleNode.get("pd");
        if (pd != null && pd.isObject()) {
            style.setIndentLevel(pd.path("l").asInt(0));
        }

        // 数字格式
        JsonNode numFmt = styleNode.get("n");
        if (numFmt != null && numFmt.isObject() && numFmt.has("pattern")) {
            String pattern = numFmt.get("pattern").asText("");
            if (!pattern.isEmpty() && !"General".equalsIgnoreCase(pattern)) {
                style.setCustom(pattern);
            }
        }

        // 边框
        JsonNode borderData = styleNode.get("bd");
        if (borderData != null && borderData.isObject()) {
            applyBorder(style, borderData, "t", BorderType.TOP_BORDER);
            applyBorder(style, borderData, "b", BorderType.BOTTOM_BORDER);
            applyBorder(style, borderData, "l", BorderType.LEFT_BORDER);
            applyBorder(style, borderData, "r", BorderType.RIGHT_BORDER);
            applyBorder(style, borderData, "tl_br", BorderType.DIAGONAL_DOWN);
            applyBorder(style, borderData, "bl_tr", BorderType.DIAGONAL_UP);
        }
    }

    private static void applyBorder(Style style, JsonNode borderData, String sideKey, int borderType) {
        JsonNode side = borderData.get(sideKey);
        if (side == null || side.isNull() || !side.isObject()) {
            // 该侧没有边框定义,清空边框
            style.setBorder(borderType, CellBorderType.NONE, Color.getBlack());
            return;
        }

        int lineStyle = side.path("s").asInt(1);
        if (lineStyle < CellBorderType.NONE || lineStyle > CellBorderType.SLANTED_DASH_DOT) {
            lineStyle = CellBorderType.THIN;
        }

        Color color = null;
        JsonNode cl = side.get("cl");
        if (cl != null && cl.isObject()) {
            color = parseColor(cl.path("rgb").asText(""));
        }
        if (color == null) {
            color = Color.getBlack();
        }
        style.setBorder(borderType, lineStyle, color);
    }

    private static void applyFontStyle(Font font, JsonNode styleNode) {
        if (styleNode.has("ff") && styleNode.get("ff").isTextual()) {
            font.setName(styleNode.get("ff").asText());
        }
        if (styleNode.has("fs") && styleNode.get("fs").isNumber()) {
            font.setDoubleSize(styleNode.get("fs").asDouble());
        }
        if (styleNode.has("bl") && styleNode.get("bl").isValueNode()) {
            font.setBold(styleNode.get("bl").asInt(0) != 0);
        }
        if (styleNode.has("it") && styleNode.get("it").isValueNode()) {
            font.setItalic(styleNode.get("it").asInt(0) != 0);
        }

        JsonNode ul = styleNode.get("ul");
        if (ul != null && ul.isObject() && ul.path("s").asInt(0) != 0) {
            int underlineType = ul.path("t").asInt(0);
            font.setUnderline(underlineType == 1 ? FontUnderlineType.DOUBLE : FontUnderlineType.SINGLE);
        }

        JsonNode st = styleNode.get("st");
        if (st != null && st.isObject()) {
            font.setStrikeout(st.path("s").asInt(0) != 0);
        }

        JsonNode cl = styleNode.get("cl");
        if (cl != null) {
            Color color = colorFromNode(cl);
            if (color != null) {
                font.setColor(color);
            }
        }

        if (styleNode.has("va")) {
            int baseline = styleNode.get("va").asInt(1);
            font.setSubscript(baseline == 2);
            font.setSuperscript(baseline == 3);
        }
    }

    private static int mapHorizontalAlignment(int univerValue) {
        switch (univerValue) {
            case 1:
                return TextAlignmentType.LEFT;
            case 2:
                return TextAlignmentType.CENTER;
            case 3:
                return TextAlignmentType.RIGHT;
            default:
                return TextAlignmentType.GENERAL;
        }
    }

    private static int mapVerticalAlignment(int univerValue) {
        switch (univerValue) {
            case 1:
                return TextAlignmentType.TOP;
            case 2:
                return TextAlignmentType.CENTER;
            case 3:
                return TextAlignmentType.BOTTOM;
            default:
                return TextAlignmentType.GENERAL;
        }
    }

    // ------------------------------------------------------------------
    // Resources (自动筛选等)
    // ------------------------------------------------------------------

    private static void applyResources(Workbook workbook, JsonNode resources, Map<String, Worksheet> sheetMap) {
        if (!resources.isArray()) {
            return;
        }
        for (JsonNode resource : resources) {
            if (!resource.isObject()) {
                continue;
            }
            String name = resource.path("name").asText("");
            String data = resource.path("data").asText("");

            if ("SHEET_FILTER_PLUGIN".equals(name) && !data.isEmpty()) {
                applyFilterResource(data, sheetMap);
            } else if ("SHEET_NOTE_PLUGIN".equals(name) && !data.isEmpty()) {
                applyNoteResource(data, sheetMap);
            } else if ("SHEET_DRAWING_PLUGIN".equals(name) && !data.isEmpty()) {
                applyDrawingResource(data, sheetMap);
            } else if ("SHEET_CONDITIONAL_FORMATTING_PLUGIN".equals(name) && !data.isEmpty()) {
                applyConditionalFormattingResource(data, sheetMap);
            } else if ("SHEET_PRINT_AREA_PLUGIN".equals(name) && !data.isEmpty()) {
                applyPrintAreaResource(data, sheetMap);
            }
        }
    }

    private static void applyNoteResource(String data, Map<String, Worksheet> sheetMap) {
        try {
            JsonNode noteData = MAPPER.readTree(data);
            if (!noteData.isObject()) {
                return;
            }
            Iterator<Map.Entry<String, JsonNode>> sheetIt = noteData.fields();
            while (sheetIt.hasNext()) {
                Map.Entry<String, JsonNode> sheetEntry = sheetIt.next();
                Worksheet worksheet = sheetMap.get(sheetEntry.getKey());
                if (worksheet == null || !sheetEntry.getValue().isObject()) {
                    continue;
                }
                Iterator<Map.Entry<String, JsonNode>> rowIt = sheetEntry.getValue().fields();
                while (rowIt.hasNext()) {
                    Map.Entry<String, JsonNode> rowEntry = rowIt.next();
                    int row = parseInt(rowEntry.getKey(), -1);
                    if (row < 0 || !rowEntry.getValue().isObject()) {
                        continue;
                    }
                    Iterator<Map.Entry<String, JsonNode>> colIt = rowEntry.getValue().fields();
                    while (colIt.hasNext()) {
                        Map.Entry<String, JsonNode> colEntry = colIt.next();
                        int col = parseInt(colEntry.getKey(), -1);
                        if (col < 0 || !colEntry.getValue().isObject()) {
                            continue;
                        }
                        JsonNode noteNode = colEntry.getValue();
                        int commentIndex = worksheet.getComments().add(row, col);
                        Comment comment = worksheet.getComments().get(commentIndex);
                        comment.setNote(noteNode.path("note").asText(""));
                        comment.setVisible(noteNode.path("show").asBoolean(true));
                        if (noteNode.has("width") && noteNode.get("width").isNumber()) {
                            comment.setWidth(noteNode.get("width").asInt());
                        }
                        if (noteNode.has("height") && noteNode.get("height").isNumber()) {
                            comment.setHeight(noteNode.get("height").asInt());
                        }
                    }
                }
            }
        } catch (IOException e) {
            // 忽略无法解析的插件数据
        }
    }

    private static void applyDrawingResource(String data, Map<String, Worksheet> sheetMap) {
        try {
            JsonNode drawingData = MAPPER.readTree(data);
            if (!drawingData.isObject()) {
                return;
            }
            Iterator<Map.Entry<String, JsonNode>> sheetIt = drawingData.fields();
            while (sheetIt.hasNext()) {
                Map.Entry<String, JsonNode> sheetEntry = sheetIt.next();
                Worksheet worksheet = sheetMap.get(sheetEntry.getKey());
                if (worksheet == null) {
                    continue;
                }
                JsonNode drawings = sheetEntry.getValue().path("data");
                if (!drawings.isObject()) {
                    continue;
                }
                Iterator<Map.Entry<String, JsonNode>> drawingIt = drawings.fields();
                while (drawingIt.hasNext()) {
                    Map.Entry<String, JsonNode> drawingEntry = drawingIt.next();
                    JsonNode drawing = drawingEntry.getValue();
                    if (!drawing.isObject()) {
                        continue;
                    }
                    String source = drawing.path("source").asText("");
                    int commaIndex = source.indexOf(',');
                    if (!source.startsWith("data:") || commaIndex < 0) {
                        continue;
                    }
                    byte[] imageBytes;
                    try {
                        imageBytes = Base64.getDecoder().decode(source.substring(commaIndex + 1));
                    } catch (IllegalArgumentException e) {
                        continue;
                    }
                    JsonNode transform = drawing.path("sheetTransform");
                    if (!transform.isObject()) {
                        transform = drawing.path("axisAlignSheetTransform");
                    }
                    JsonNode from = transform.path("from");
                    int row = from.path("row").asInt(0);
                    int col = from.path("column").asInt(0);
                    try (ByteArrayInputStream imageIn = new ByteArrayInputStream(imageBytes)) {
                        worksheet.getPictures().add(row, col, imageIn);
                    }
                }
            }
        } catch (Exception e) {
            // 单个图片失败时不影响整个工作簿生成
        }
    }

    private static void applyConditionalFormattingResource(String data, Map<String, Worksheet> sheetMap) {
        try {
            JsonNode cfData = MAPPER.readTree(data);
            if (!cfData.isObject()) {
                return;
            }
            Iterator<Map.Entry<String, JsonNode>> sheetIt = cfData.fields();
            while (sheetIt.hasNext()) {
                Map.Entry<String, JsonNode> sheetEntry = sheetIt.next();
                Worksheet worksheet = sheetMap.get(sheetEntry.getKey());
                if (worksheet == null || !sheetEntry.getValue().isArray()) {
                    continue;
                }
                for (JsonNode cf : sheetEntry.getValue()) {
                    if (cf.isObject()) {
                        applyConditionalFormatting(worksheet, cf);
                    }
                }
            }
        } catch (IOException e) {
            // 忽略无法解析的插件数据
        }
    }

    private static void applyConditionalFormatting(Worksheet worksheet, JsonNode cfNode) {
        JsonNode rule = cfNode.get("rule");
        JsonNode ranges = cfNode.path("ranges");
        if (rule == null || !rule.isObject() || !ranges.isArray()) {
            return;
        }

        List<CellArea> areas = new ArrayList<>();
        for (JsonNode range : ranges) {
            int startRow = range.path("startRow").asInt(-1);
            int startCol = range.path("startColumn").asInt(-1);
            int endRow = range.path("endRow").asInt(-1);
            int endCol = range.path("endColumn").asInt(-1);
            if (startRow < 0 || startCol < 0 || endRow < startRow || endCol < startCol) {
                continue;
            }
            areas.add(CellArea.createCellArea(startRow, startCol, endRow, endCol));
        }
        if (areas.isEmpty()) {
            return;
        }

        int cfIndex = worksheet.getConditionalFormattings().add();
        FormatConditionCollection fcc = worksheet.getConditionalFormattings().get(cfIndex);
        for (CellArea area : areas) {
            fcc.addArea(area);
        }

        String type = rule.path("type").asText("");
        if ("highlightCell".equals(type)) {
            addHighlightCellCondition(fcc, cfNode, rule);
        } else if ("colorScale".equals(type)) {
            addColorScaleCondition(fcc, cfNode, rule);
        } else if ("dataBar".equals(type)) {
            addDataBarCondition(fcc, cfNode, rule);
        } else if ("iconSet".equals(type)) {
            addIconSetCondition(fcc, cfNode, rule);
        }
    }

    private static void addHighlightCellCondition(FormatConditionCollection fcc, JsonNode cfNode, JsonNode rule) {
        String operator = rule.path("operator").asText("equal");
        String formula1 = toConditionFormulaValue(rule.path("value"));
        int op = mapHighlightCellOperator(operator);

        int idx = fcc.addCondition(FormatConditionType.CELL_VALUE);
        FormatCondition condition = fcc.get(idx);
        condition.setOperator(op);
        condition.setFormula1(formula1);
        applyConditionStyle(condition, rule.path("style"));
        configureCondition(condition, cfNode);
    }

    private static void addColorScaleCondition(FormatConditionCollection fcc, JsonNode cfNode, JsonNode rule) {
        JsonNode config = rule.path("config");
        if (!config.isArray() || config.size() == 0) {
            return;
        }
        int idx = fcc.addCondition(FormatConditionType.COLOR_SCALE);
        FormatCondition condition = fcc.get(idx);
        ColorScale colorScale = condition.getColorScale();
        colorScale.setIs3ColorScale(config.size() >= 3);

        if (config.size() >= 1) {
            applyColorScaleItem(colorScale.getMinCfvo(), config.get(0));
            Color color = colorFromNode(config.get(0).path("color"));
            if (color != null) {
                colorScale.setMinColor(color);
            }
        }
        if (config.size() >= 2) {
            applyColorScaleItem(colorScale.getMaxCfvo(), config.get(1));
            Color color = colorFromNode(config.get(1).path("color"));
            if (color != null) {
                colorScale.setMaxColor(color);
            }
        }
        if (config.size() >= 3) {
            applyColorScaleItem(colorScale.getMidCfvo(), config.get(1));
            Color midColor = colorFromNode(config.get(1).path("color"));
            if (midColor != null) {
                colorScale.setMidColor(midColor);
            }
            applyColorScaleItem(colorScale.getMaxCfvo(), config.get(2));
            Color maxColor = colorFromNode(config.get(2).path("color"));
            if (maxColor != null) {
                colorScale.setMaxColor(maxColor);
            }
        }

        configureCondition(condition, cfNode);
    }

    private static void addDataBarCondition(FormatConditionCollection fcc, JsonNode cfNode, JsonNode rule) {
        JsonNode config = rule.path("config");
        if (!config.isObject()) {
            return;
        }
        int idx = fcc.addCondition(FormatConditionType.DATA_BAR);
        FormatCondition condition = fcc.get(idx);
        DataBar dataBar = condition.getDataBar();
        dataBar.setShowValue(rule.path("isShowValue").asBoolean(true));
        dataBar.setBarFillType(config.path("isGradient").asBoolean(true)
                ? DataBarFillType.GRADIENT : DataBarFillType.SOLID);

        if (config.has("min") && !config.get("min").isNull()) {
            applyColorScaleItem(dataBar.getMinCfvo(), config.get("min"));
        }
        if (config.has("max") && !config.get("max").isNull()) {
            applyColorScaleItem(dataBar.getMaxCfvo(), config.get("max"));
        }

        Color positiveColor = colorFromNode(config.get("positiveColor"));
        if (positiveColor != null) {
            dataBar.setColor(positiveColor);
        }
        Color negativeColor = colorFromNode(config.get("nativeColor"));
        if (negativeColor != null) {
            dataBar.getNegativeBarFormat().setColorType(DataBarNegativeColorType.COLOR);
            dataBar.getNegativeBarFormat().setColor(negativeColor);
        }

        configureCondition(condition, cfNode);
    }

    private static void addIconSetCondition(FormatConditionCollection fcc, JsonNode cfNode, JsonNode rule) {
        JsonNode config = rule.path("config");
        if (!config.isArray() || config.size() == 0) {
            return;
        }

        String iconTypeName = iconSetTypeName(config.get(0).get("iconType"));
        int iconSetType = mapIconSetType(iconTypeName);
        if (iconSetType == IconSetType.NONE) {
            return;
        }

        int idx = fcc.addCondition(FormatConditionType.ICON_SET);
        FormatCondition condition = fcc.get(idx);
        IconSet iconSet = condition.getIconSet();
        iconSet.setType(iconSetType);
        iconSet.setShowValue(rule.path("isShowValue").asBoolean(true));

        // Univer 的 icon config 按“图标展示顺序”排列（如 3Arrows: 绿色→黄色→红色），
        // Aspose.Cells 的 cfvo 从低到高排列，因此需要逆序填充。
        int total = iconSet.getCfvos().getCount();
        int count = Math.min(total, config.size());
        for (int i = 0; i < count; i++) {
            applyIconSetItem(iconSet.getCfvos().get(total - 1 - i), config.get(i));
        }

        configureCondition(condition, cfNode);
    }

    private static void applyIconSetItem(ConditionalFormattingValue cfvo, JsonNode item) {
        applyColorScaleItem(cfvo, item.path("value"));
        String operator = item.path("operator").asText("");
        // greaterThan/lessThan 表示严格比较；其余（greaterThanOrEqual、lessThanOrEqual、equal 等）视为 >=。
        boolean gte = !("greaterThan".equals(operator) || "lessThan".equals(operator)
                || "notEqual".equals(operator) || "notBetween".equals(operator) || "between".equals(operator));
        cfvo.setGTE(gte);
    }

    private static String iconSetTypeName(JsonNode iconType) {
        if (iconType == null || iconType.isNull()) {
            return "";
        }
        if (iconType.isTextual()) {
            return iconType.asText("");
        }
        if (iconType.isObject() && iconType.get("value") != null && iconType.get("value").isTextual()) {
            return iconType.get("value").asText("");
        }
        return "";
    }

    private static int mapIconSetType(String iconType) {
        switch (iconType) {
            case "3Arrows": return IconSetType.ARROWS_3;
            case "3ArrowsGray": return IconSetType.ARROWS_GRAY_3;
            case "3Flags": return IconSetType.FLAGS_3;
            case "3Signs": return IconSetType.SIGNS_3;
            case "3Symbols": return IconSetType.SYMBOLS_3;
            case "3Symbols2": return IconSetType.SYMBOLS_32;
            case "3Triangles": return IconSetType.TRIANGLES_3;
            case "3TrafficLights1": return IconSetType.TRAFFIC_LIGHTS_31;
            case "3TrafficLights2": return IconSetType.TRAFFIC_LIGHTS_32;
            case "4Arrows": return IconSetType.ARROWS_4;
            case "4ArrowsGray": return IconSetType.ARROWS_GRAY_4;
            case "4Rating": return IconSetType.RATING_4;
            case "4RedToBlack": return IconSetType.RED_TO_BLACK_4;
            case "4TrafficLights": return IconSetType.TRAFFIC_LIGHTS_4;
            case "5Arrows": return IconSetType.ARROWS_5;
            case "5ArrowsGray": return IconSetType.ARROWS_GRAY_5;
            case "5Quarters": return IconSetType.QUARTERS_5;
            case "5Rating": return IconSetType.RATING_5;
            case "5Boxes": return IconSetType.BOXES_5;
            case "3Stars": return IconSetType.STARS_3;
            case "EMPTY_ICON_TYPE":
            default:
                return IconSetType.NONE;
        }
    }

    private static void applyColorScaleItem(ConditionalFormattingValue cfvo, JsonNode item) {
        if (item == null || item.isNull()) {
            return;
        }
        // 兼容两种写法：{ type: "num", value: 20 } 或直接写数值/文本。
        JsonNode valueConfig = item.isObject() ? item.get("value") : item;
        if (valueConfig == null || !valueConfig.isObject()) {
            cfvo.setType(FormatConditionValueType.NUMBER);
            if (valueConfig != null && !valueConfig.isNull()) {
                if (valueConfig.isNumber()) {
                    cfvo.setValue(valueConfig.asDouble());
                } else {
                    cfvo.setValue(valueConfig.asText());
                }
            }
            return;
        }
        String type = valueConfig.path("type").asText("");
        cfvo.setType(mapConditionValueType(type));
        JsonNode valueNode = valueConfig.get("value");
        if (valueNode != null && !valueNode.isNull()) {
            if (valueNode.isNumber()) {
                cfvo.setValue(valueNode.asDouble());
            } else {
                cfvo.setValue(valueNode.asText());
            }
        }
    }

    private static int mapConditionValueType(String type) {
        switch (type) {
            case "min":
                return FormatConditionValueType.MIN;
            case "max":
                return FormatConditionValueType.MAX;
            case "percent":
                return FormatConditionValueType.PERCENT;
            case "percentile":
                return FormatConditionValueType.PERCENTILE;
            case "formula":
                return FormatConditionValueType.FORMULA;
            case "num":
            case "number":
            default:
                return FormatConditionValueType.NUMBER;
        }
    }

    private static void configureCondition(FormatCondition condition, JsonNode cfNode) {
        condition.setStopIfTrue(cfNode.path("stopIfTrue").asBoolean(false));
        if (cfNode.has("order") && cfNode.get("order").isNumber()) {
            condition.setPriority(cfNode.get("order").asInt(1));
        }
    }

    private static void applyConditionStyle(FormatCondition condition, JsonNode styleNode) {
        if (styleNode == null || !styleNode.isObject()) {
            return;
        }
        Style style = condition.getStyle();
        applyStyle(style, styleNode);
        condition.setStyle(style);
    }

    private static int mapHighlightCellOperator(String operator) {
        switch (operator) {
            case "notEqual":
                return OperatorType.NOT_EQUAL;
            case "greaterThan":
                return OperatorType.GREATER_THAN;
            case "greaterThanOrEqual":
                return OperatorType.GREATER_OR_EQUAL;
            case "lessThan":
                return OperatorType.LESS_THAN;
            case "lessThanOrEqual":
                return OperatorType.LESS_OR_EQUAL;
            case "between":
                return OperatorType.BETWEEN;
            case "notBetween":
                return OperatorType.NOT_BETWEEN;
            case "equal":
            default:
                return OperatorType.EQUAL;
        }
    }

    private static String toConditionFormulaValue(JsonNode valueNode) {
        if (valueNode == null || valueNode.isNull()) {
            return "";
        }
        String raw = valueNode.asText("");
        String value = unescapeHtmlEntities(raw).trim();
        if (value.length() >= 2 && value.startsWith("\"") && value.endsWith("\"")) {
            // Univer 中 &quot;xxx&quot; 表示文本字面量，Aspose.Cells 会自动为文本补上引号，
            // 因此这里只保留原始文本，不要再额外加引号。
            value = value.substring(1, value.length() - 1).trim();
        }
        return value;
    }

    private static String unescapeHtmlEntities(String s) {
        return s.replace("&quot;", "\"")
                .replace("&amp;", "&")
                .replace("&lt;", "<")
                .replace("&gt;", ">")
                .replace("&#39;", "'");
    }

    private static void applyPrintAreaResource(String data, Map<String, Worksheet> sheetMap) {
        try {
            JsonNode printAreaData = MAPPER.readTree(data);
            if (!printAreaData.isObject()) {
                return;
            }
            Iterator<Map.Entry<String, JsonNode>> sheetIt = printAreaData.fields();
            while (sheetIt.hasNext()) {
                Map.Entry<String, JsonNode> sheetEntry = sheetIt.next();
                Worksheet worksheet = sheetMap.get(sheetEntry.getKey());
                if (worksheet == null) {
                    continue;
                }
                JsonNode sheetValue = sheetEntry.getValue();
                if (sheetValue.isTextual()) {
                    worksheet.getPageSetup().setPrintArea(sheetValue.asText());
                    continue;
                }
                if (!sheetValue.isObject()) {
                    continue;
                }
                JsonNode ranges = sheetValue.path("ranges");
                if (!ranges.isArray()) {
                    continue;
                }
                List<String> areas = new ArrayList<>();
                for (JsonNode range : ranges) {
                    int startRow = range.path("startRow").asInt(-1);
                    int startCol = range.path("startColumn").asInt(-1);
                    int endRow = range.path("endRow").asInt(-1);
                    int endCol = range.path("endColumn").asInt(-1);
                    if (startRow < 0 || startCol < 0 || endRow < startRow || endCol < startCol) {
                        continue;
                    }
                    areas.add(excelRange(startRow, startCol, endRow, endCol));
                }
                if (!areas.isEmpty()) {
                    worksheet.getPageSetup().setPrintArea(String.join(",", areas));
                }
            }
        } catch (IOException e) {
            // 忽略无法解析的插件数据
        }
    }

    private static void applyFilterResource(String data, Map<String, Worksheet> sheetMap) {
        try {
            JsonNode filterData = MAPPER.readTree(data);
            if (!filterData.isObject()) {
                return;
            }
            Iterator<Map.Entry<String, JsonNode>> it = filterData.fields();
            while (it.hasNext()) {
                Map.Entry<String, JsonNode> entry = it.next();
                Worksheet worksheet = sheetMap.get(entry.getKey());
                if (worksheet == null) {
                    continue;
                }
                JsonNode ref = entry.getValue().path("ref");
                if (!ref.isObject()) {
                    continue;
                }
                int startRow = ref.path("startRow").asInt(-1);
                int startCol = ref.path("startColumn").asInt(-1);
                int endRow = ref.path("endRow").asInt(-1);
                int endCol = ref.path("endColumn").asInt(-1);
                if (startRow < 0 || startCol < 0 || endRow < startRow || endCol < startCol) {
                    continue;
                }
                String range = excelRange(startRow, startCol, endRow, endCol);
                worksheet.getAutoFilter().setRange(range);
            }
        } catch (IOException e) {
            // 忽略无法解析的插件数据
        }
    }

    private static String excelRange(int startRow, int startCol, int endRow, int endCol) {
        return columnName(startCol) + (startRow + 1) + ":" + columnName(endCol) + (endRow + 1);
    }

    private static String columnName(int col) {
        StringBuilder sb = new StringBuilder();
        int c = col;
        while (c >= 0) {
            sb.insert(0, (char) ('A' + c % 26));
            c = c / 26 - 1;
        }
        return sb.toString();
    }

    // ------------------------------------------------------------------
    // Utilities
    // ------------------------------------------------------------------

    private static int parseInt(String value, int defaultValue) {
        if (value == null) {
            return defaultValue;
        }
        try {
            return Integer.parseInt(value.trim());
        } catch (NumberFormatException e) {
            return defaultValue;
        }
    }

    /**
     * 从 Univer 颜色节点中取色。可能是字符串 "#FF7128"，也可能是 {"rgb":"#FF7128"}。
     */
    private static Color colorFromNode(JsonNode colorNode) {
        if (colorNode == null || colorNode.isNull()) {
            return null;
        }
        String rgb = colorNode.isTextual() ? colorNode.asText("") : colorNode.path("rgb").asText("");
        return parseColor(rgb);
    }

    /**
     * 解析 #RRGGBB / #RGB。
     */
    private static Color parseColor(String rgb) {
        if (rgb == null || rgb.isEmpty()) {
            return null;
        }
        String s = rgb.trim();
        if (s.startsWith("#")) {
            s = s.substring(1);
        }
        if (s.length() == 3) {
            char[] chars = new char[6];
            chars[0] = chars[1] = s.charAt(0);
            chars[2] = chars[3] = s.charAt(1);
            chars[4] = chars[5] = s.charAt(2);
            s = new String(chars);
        }
        if (s.length() != 6) {
            return null;
        }
        try {
            int value = Integer.parseInt(s, 16);
            int r = (value >> 16) & 0xFF;
            int g = (value >> 8) & 0xFF;
            int b = value & 0xFF;
            return Color.fromArgb(255, r, g, b);
        } catch (NumberFormatException e) {
            return null;
        }
    }
}
