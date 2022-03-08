package com.tkachenko.buyerhelper.utils;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.GregorianCalendar;

public class DateUtils {
    public static String getYear (GregorianCalendar calendar) {
        DateFormat yearFormat = new SimpleDateFormat("yyyy");
        return yearFormat.format(calendar.getTime());
    }

    public static String getMonth (GregorianCalendar calendar) {
        DateFormat monthFormat = new SimpleDateFormat("MM");
        return monthFormat.format(calendar.getTime());
    }

    public static String getDay (GregorianCalendar calendar) {
        DateFormat dayFormat = new SimpleDateFormat("dd");
        return dayFormat.format(calendar.getTime());
    }

    public static String getTime (GregorianCalendar calendar) {
        DateFormat timeFormat = new SimpleDateFormat("HH-mm");
        return timeFormat.format(calendar.getTime());
    }
}
