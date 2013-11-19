package com.ecfront.easybi.excelconverter.inner.util;

/**
 * Unit Convert Helper
 *
 * @author gudaoxuri
 */
public class UnitConvertHelper {

    /**
     * Convert 26 TO 10 ,Use word TO number
     * @param s
     * @return
     */
    public static int convert26to10(String s) {
        if (null != s && !"".equals(s)) {
            int n = 0;
            char[] tmp = s.toCharArray();
            for (int i = tmp.length - 1, j = 1; i >= 0; i--, j *= 26) {
                char c = Character.toUpperCase(tmp[i]);
                if (c < 'A' || c > 'Z') return 0;
                n += ((int) c - 64) * j;
            }
            return n;
        }
        return 0;
    }
}
