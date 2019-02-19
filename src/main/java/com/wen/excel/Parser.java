package com.wen.excel;

import java.io.File;
import java.util.Map;

/**
 * @author admin
 * @date 2019-2-18 17:28
 */
public interface Parser {

    Map<String,Object> parser(File file) throws Exception;

}
