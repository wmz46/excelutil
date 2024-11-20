package com.iceolive.util;

import java.util.regex.Pattern;

public class ZeroWidthCharUtil {
    private static final Pattern ZERO_WIDTH_CHAR_PATTERN = Pattern.compile("[\\u200B-\\u200D\\u2060\\uFEFF]");

    private ZeroWidthCharUtil(){

    }
    public static String filterZeroWidthChars(String input) {
        if (input == null) {
            return null;
        }
        return ZERO_WIDTH_CHAR_PATTERN.matcher(input).replaceAll("");
    }
}
