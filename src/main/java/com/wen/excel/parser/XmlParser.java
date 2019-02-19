package com.wen.excel.parser;

import com.wen.excel.Parser;
import org.jdom2.Document;
import org.jdom2.Element;
import org.jdom2.input.SAXBuilder;

import java.io.File;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author Echon
 * @date 2019-2-18 15:56
 */
public class XmlParser extends AbstractParser implements Parser {

    public Map<String, Object> parser(File file) throws Exception {
        Map<String, Object> map = new HashMap<String, Object>();
        SAXBuilder sb = new SAXBuilder();
        Document document = sb.build(file);
        Element root = document.getRootElement();

        map.put("table", eleAttrValueTrim(root, "table"));
        map.put("alias", eleAttrValueTrim(root, "alias"));
        map.put("fields", readFields(root.getChild("fields")));
        return map;
    }

    public Map<String, Map<String, Object>> readFields(Element eleFields) {
        if (eleFields == null) return null;

        Map<String, Map<String, Object>> fields = new HashMap<String, Map<String, Object>>();
        List<Element> childes = eleFields.getChildren();
        Map<String, Object> field;
        for (Element children : childes) {
            field = new HashMap<>();
            field.put("alias", eleAttrValueTrim(children, "alias"));
            field.put("type", eleAttrValueTrim(children, "type"));
            field.put("name", eleAttrValueTrim(children, "name"));
            field.put("domain", eleAttrValueTrim(children, "domain"));
            field.put("filter", eleAttrValueTrim(children, "filter"));
            field.put("dictionary", eleAttrValueTrim(children, "dictionary"));
            String exSort = eleAttrValueTrim(children, "exSort");
            if (exSort != null) {
                field.put("exSort", Integer.valueOf(exSort));
            }
            field.put("exAlias", eleAttrValueTrim(children, "exAlias"));
            String imSort = eleAttrValueTrim(children, "imSort");
            if (imSort != null) {
                field.put("imSort", Integer.valueOf(imSort));
            }

            fields.put(eleAttrValueTrim(children, "name"), field);
        }

        return fields;
    }
}
