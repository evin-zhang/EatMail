package org.shinear;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.Map;

class UtilsTest {

    String t = "安装完成时间：11月16日\n" +
            "客户姓名:   匡捷\n" +
            "车辆品牌：比亚迪\n" +
            "增项价格： 0\n" +
            "增项内容:    无\n" +
            "物料用量:   3*6  12米\n" +
            "联塑pvc25  15米\n" +
            "安装人员:程新  陈明祥\n" +
            "出发地:麻城";
    String p = "3*6  12米\n" +
            "联塑pvc25  15米";

    @Test
    void TestC() {
        Map<String, String> result = new LinkedHashMap<>();
        Arrays.stream(t.split("\n")).forEach(line -> {
            if (line.contains("：") || line.contains(":")) {
                String[] parts = line.split("[：:]", 2);
                result.put(parts[0].trim(), parts[1].trim());
            } else {
                // Append the line to the last key if there's no "：" in the line
                result.computeIfPresent(result.keySet().stream().reduce((first, second) -> second).orElse(null), (key, value) -> value + "\n" + line.trim());
            }
        });

        Assertions.assertTrue(result.containsKey("物料用量"));
        Assertions.assertEquals(p, result.get("物料用量"));
    }
}
