package com.test;

import java.io.File;

/**
 * 简单示例：把项目内的 univerData.json 转成 Excel 或 PDF。
 */
public class Main {

    public static void main(String[] args) throws Exception {
        String input = args.length > 0 ? args[0] : "src/main/resources/data/univerData2.json";
        String output = args.length > 1 ? args[1] : "target/univer-output2.pdf";

        if (output.toLowerCase(java.util.Locale.ROOT).endsWith(".pdf")) {
            UniverJson2Excel.convertToPdf(new File(input), output);
            System.out.println("已生成 PDF: " + new File(output).getAbsolutePath());
        } else {
            UniverJson2Excel.convertToFile(new File(input), output);
            System.out.println("已生成 Excel: " + new File(output).getAbsolutePath());
        }
    }
}
