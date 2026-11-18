package com.test;

import java.io.File;

/**
 * 简单示例：把项目内的 univerData.json 转成 Excel。
 */
public class Main {

    public static void main(String[] args) throws Exception {
        String input = args.length > 0 ? args[0] : "src/main/resources/data/univerData.json";
        String output = args.length > 1 ? args[1] : "target/univer-output.xlsx";

        UniverJson2Excel.convertToFile(new File(input), output);
        System.out.println("已生成: " + new File(output).getAbsolutePath());
    }
}
