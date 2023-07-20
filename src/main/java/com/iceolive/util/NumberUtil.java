package com.iceolive.util;

import org.apache.commons.lang3.StringUtils;

/**
 * @author wangmianzhe
 */
public class NumberUtil {
    public static int compare(Object o1, Object o2) {
        if (o1 == null || o1 == null) {
            throw new IllegalArgumentException("比较参数不能为null");
        }
        if(!StringUtils.isNumeric(String.valueOf(o1)) || !StringUtils.isNumeric(String.valueOf(o2))){
            throw new IllegalArgumentException("比较参数必须是数值或数值字符串");
        }
        Double d1 = StringUtil.parse(String.valueOf(o1), Double.class);
        Double d2 = StringUtil.parse(String.valueOf(o2), Double.class);
        return d1.compareTo(d2);
    }

    public static boolean equals(Object o1, Object o2) {
        if (o1 == null && o1 == null) {
            return true;
        } else if (o1 == null) {
            return false;
        } else if (o2 == null) {
            return false;
        } else {
            return compare(o1, o2) == 0;
        }
    }

    public static boolean lessOrEqual(Object o1, Object o2) {
        return compare(o1, o2) <= 0;
    }

    public static boolean greaterOrEqual(Object o1, Object o2) {
        return compare(o1, o2) >= 0;
    }

}
