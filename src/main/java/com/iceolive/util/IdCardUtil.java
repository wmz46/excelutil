package com.iceolive.util;

import java.util.Date;

public class IdCardUtil {
    public static boolean validate(String idCard) {
        if (idCard == null || idCard.length() != 18) {
            return false;
        }
        String birthday = idCard.substring(6, 14);

        try {
            String format = "yyyyMMdd";
            if (!StringUtil.format(StringUtil.parse(birthday, format, Date.class), format).equals(birthday)) {
                return false;
            }
        } catch (Exception e) {
            return false;
        }
        char[] digits = idCard.toCharArray();
        int sum = 0;
        int[] weight = new int[]{7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2};
        char[] checkDigits = new char[]{'1', '0', 'X', '9', '8', '7', '6', '5', '4', '3', '2'};

        for (int i = 0; i < 17; i++) {
            if (!Character.isDigit(digits[i])) {
                return false;
            }
            sum += (digits[i] - '0') * weight[i];
        }

        int remainder = sum % 11;
        char checkDigit = checkDigits[remainder];

        return checkDigit == digits[17];
    }
}
