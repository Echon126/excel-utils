package com.wen.excel.configuration;

import com.wen.excel.Parser;
import com.wen.excel.parser.XmlParser;
import com.wen.excel.parser.PropertieParser;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.*;
import java.util.concurrent.ConcurrentHashMap;

/**
 * @author admin
 * @date 2019-2-18 17:17
 */
public class BuilderConfiguration {

    public Map<String, Map<String, Object>> configeValue = new ConcurrentHashMap<String, Map<String, Object>>(64);

    public File[] files;

    public Map<String, Parser> parserMap;

    private String downDir;

    public BuilderConfiguration(String downDir, String path) {
        this.downDir = downDir;
        init(path);
    }

    public void init(String path) {
        parserMap = new HashMap<String, Parser>();
        parserMap.put("xml", new XmlParser());
        parserMap.put("properties", new PropertieParser());

        loadConfiguration(path);
        builderParser();
    }


    public void loadConfiguration(String path) {
        File xmls = new File(path);
        if (!xmls.exists() || !xmls.isDirectory()) return;
        files = xmls.listFiles();
    }

    public void builderParser() {
        for (File f : files) {
            String[] tem = f.getName().split("\\.");
            String suffix = tem[tem.length - 1];
            Parser fileParser = parserMap.get(suffix);
            try {
                configeValue.put(f.getName().replace(".xml", ""), fileParser.parser(f));
            } catch (Exception e) {
                System.out.println("构建模板文件时发生错误" + e.getMessage());
                e.printStackTrace();
            }
        }
    }

    /**
     * @param model 模板的文件名
     * @param list  需要导出的数据
     * @return
     */
    public String builderUtils(String model, List<Map<String, Object>> list) throws IOException {
        long f = new Date().getTime();
        Map<String, Object> map = configeValue.get(model);
        if (list.size() == 0 || list == null) {
            return "导出数据为空！";
        }
        String fileName = map.get("alias") + "-" + new Date().getTime() + ".xls";
        Map<String, Map<String, Object>> fields = (Map<String, Map<String, Object>>) map.get("fields");
        TreeMap<Integer, Map<String, Object>> treeMap = new TreeMap<Integer, Map<String, Object>>();
        Map<String, Integer> hashMap = new HashMap<String, Integer>();
        List<Integer> exportFieldWidth = new ArrayList<Integer>();
        Iterator<Map.Entry<String, Map<String, Object>>> iterator = fields.entrySet().iterator();
        while (iterator.hasNext()) {
            Map.Entry<String, Map<String, Object>> field = iterator.next();
            treeMap.put((Integer) field.getValue().get("exSort"), field.getValue());
            hashMap.put(field.getKey(), (Integer) field.getValue().get("exSort"));

            exportFieldWidth.add(15);
        }
        Workbook workbook = new HSSFWorkbook();
        // 生成一个表格
        Sheet sheet = workbook.createSheet(fileName);

        // 产生表格标题行
        int index = 0;
        Row row = sheet.createRow(index);
        Iterator<Map.Entry<Integer, Map<String, Object>>> itTree = treeMap.entrySet().iterator();
        while (itTree.hasNext()) {
            Map.Entry<Integer, Map<String, Object>> entryTree = itTree.next();
            Cell cell = row.createCell(entryTree.getKey());
            RichTextString text = new HSSFRichTextString((String) entryTree.getValue().get("alias"));
            cell.setCellValue(text);
        }
        //设置每行的列宽
        for (int i = 0; i < exportFieldWidth.size(); i++) {
            //256=65280/255
            sheet.setColumnWidth(i, 256 * exportFieldWidth.get(i));
        }
        // 循环插入剩下的集合
        for (Map<String, Object> dm : list) {
            // 从第二行开始写，第一行是标题
            index++;
            row = sheet.createRow(index);
            Iterator<Map.Entry<String, Object>> it = dm.entrySet().iterator();
            while (it.hasNext()) {
                Map.Entry<String, Object> entry = it.next();
                if (hashMap.containsKey(entry.getKey() + "")) {
                    Cell cell = row.createCell(hashMap.get(entry.getKey() + ""));
                    cell.setCellValue(entry.getValue().toString());
                }

            }
        }
        OutputStream out = new FileOutputStream(downDir+fileName);
        workbook.write(out);
        out.flush();
        out.close();
        long e = new Date().getTime();
        System.out.println("共" + (e - f) + "毫秒完成");
        return fileName;
    }


}
