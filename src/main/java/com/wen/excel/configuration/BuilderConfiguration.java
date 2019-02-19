package com.wen.excel.configuration;

import com.wen.excel.Parser;
import com.wen.excel.parser.PropertieParser;
import com.wen.excel.parser.XmlParser;
import com.wen.excel.util.Constants;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
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
 * @author Echon
 * @date 2019-2-18 17:17
 */

public class BuilderConfiguration {
    protected final Log logger = LogFactory.getLog(getClass());

    public Map<String, Map<String, Object>> configeValue = new ConcurrentHashMap<String, Map<String, Object>>(64);

    public File[] files;

    public Map<String, Parser> parserMap;

    private String downDir;

    public BuilderConfiguration(String downDir, String path) {
        this.downDir = downDir;
        init(path);
    }

    public void init(String path) {
        parserMap = new HashMap<>();

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
                logger.error("构建模板文件错误. 错误信息: {}", e);

            }
        }
    }

    public String builderUtils(String model, List<Map<String, Object>> list) {
        long f = new Date().getTime();
        Map<String, Object> map = configeValue.get(model);
        String fileName = map.get("alias") + "-" + new Date().getTime() + Constants.EXCEL_SUFFIX;
        Map<String, Map<String, Object>> fields = (Map<String, Map<String, Object>>) map.get("fields");
        TreeMap<Integer, Map<String, Object>> treeMap = new TreeMap<Integer, Map<String, Object>>();
        Map<String, Integer> hashMap = new HashMap<String, Integer>();
        List<Integer> exportFieldWidth = new ArrayList<Integer>();

        for (Map.Entry<String, Map<String, Object>> field : fields.entrySet()) {
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
        for (Map.Entry<Integer, Map<String, Object>> entryTree : treeMap.entrySet()) {
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
            for (Map.Entry<String, Object> entry : dm.entrySet()) {
                if (hashMap.containsKey(entry.getKey() + "")) {
                    Cell cell = row.createCell(hashMap.get(entry.getKey() + ""));
                    cell.setCellValue(entry.getValue().toString());
                }

            }
        }

        try (OutputStream out = new FileOutputStream(downDir + fileName)) {
            workbook.write(out);
        } catch (IOException e) {
            logger.error("生成excel文件时发生错误.错误信息:{}", e);
        }

        long e = new Date().getTime();
        logger.debug("共" + (e - f) + "毫秒完成");
        return fileName;
    }


}
