package com.tkachenko.buyerhelper.service.mmk;

import java.util.HashMap;
import java.util.Map;

public enum MonthSheets {
    JANUARY(1, "Январь_2022"),
    FEBRUARY(2, "Февраль_2022"),
    MARCH(3, "Март_2022"),
    APRIL(4, "Апрель_2022"),
    MAY(5, "Май_2022"),
    JUNE(6, "Июнь_2022"),
    JULY(7, "Июль_2022"),
    AUGUST(8, "Август_2022"),
    SEPTEMBER(9, "Сентябрь_2022"),
    OCTOBER(10, "Октябрь_2022"),
    NOVEMBER(11, "Ноябрь_2022"),
    DECEMBER(12, "Декабрь_2022");

    private final int intValue;
    private final String sheetName;

    MonthSheets(int intValue, String sheetName) {
        this.intValue = intValue;
        this.sheetName = sheetName;
    }

    public int getIntValue() {
        return intValue;
    }

    public String getSheetName() {
        return sheetName;
    }

    private static final Map <Integer,MonthSheets> mapIntegerMonth;
    static {
        mapIntegerMonth = new HashMap<Integer,MonthSheets>();
        for (MonthSheets monthSheet : MonthSheets.values()) {
            mapIntegerMonth.put(monthSheet.intValue, monthSheet);
        }
    }
    public static MonthSheets findByIntValue(int i) {
        return mapIntegerMonth.get(i);
    }

    private static final Map <String,MonthSheets> mapStringMonth;
    static {
        mapStringMonth = new HashMap<String,MonthSheets>();
        for (MonthSheets month : MonthSheets.values()) {
            mapStringMonth.put(month.getSheetName(), month);
        }
    }
    public static MonthSheets findBySheetName(String name) {
        return mapStringMonth.get(name);
    }
}
