package com.wen.excel.support;

import com.wen.excel.Parser;
import org.springframework.beans.BeansException;
import org.springframework.context.ApplicationContext;
import org.springframework.context.ApplicationContextAware;
import org.springframework.stereotype.Component;

import java.util.Map;

/**
 * @author admin
 * @date 2019-2-19 15:11
 */
@Component
public class ServiceLocator implements ApplicationContextAware {


    private Map<String, Parser> map;

    @Override
    public void setApplicationContext(ApplicationContext applicationContext) throws BeansException {
        Map<String, Parser> map = applicationContext.getBeansOfType(Parser.class);
    }


    public Map<String, Parser> getMap() {
        return map;
    }
}
