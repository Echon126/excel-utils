package com.wen.excel.parser;

import com.wen.excel.Parser;
import org.apache.commons.lang3.StringUtils;
import org.jdom2.Element;

/**
 * @author admin
 * @date 2019-2-18 18:01
 */
public abstract class AbstractParser implements Parser {

    public String eleText(Element e) {
        return e == null ? null : e.getTextTrim();
    }

    public String eleAttrValueTrim(Element e, String attr, String dvalue) {
        String v = eleAttrValueTrim(e, attr);
        return v == null ? dvalue : v;
    }

    public String eleAttrValueTrim(Element e, String attr) {
        if (e == null) {
            return null;
        }

        String value = e.getAttributeValue(attr);

        if (StringUtils.isEmpty(value) || value.trim().length() == 0) {
            return null;
        }
        return value.trim();
    }

}
