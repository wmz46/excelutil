package com.iceolive.util;

import java.beans.BeanInfo;
import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.lang.reflect.Method;
import java.util.HashMap;
import java.util.Map;

public class BeanUtil {
    private BeanUtil() {
    }
    /**
     * 扁平化第一层属性（不递归处理嵌套对象）
     *
     * @param bean 目标对象
     * @return 键为属性名，值为原始对象（包括嵌套对象）
     */
    public static Map<String, Object> beanToMapShallow(Object bean) {
        Map<String, Object> map = new HashMap<>(16);
        if (bean == null) {
            return map;
        }
        if(bean instanceof Map){
            return (Map<String, Object>) bean;
        }

        try {
            BeanInfo beanInfo = Introspector.getBeanInfo(bean.getClass(),  Object.class);
            for (PropertyDescriptor pd : beanInfo.getPropertyDescriptors())  {
                String name = pd.getName();
                if ("class".equals(name)) {
                    continue; // 排除class属性
                }

                Method getter = pd.getReadMethod();
                if (getter != null) {
                    // 直接获取原始对象，不递归处理
                    Object value = getter.invoke(bean);
                    map.put(name,  value);
                }
            }
        } catch (Exception e) {
            throw new RuntimeException("浅层转换失败: " + e.getMessage());
        }
        return map;
    }
}
