package com.tkachenko.buyerhelper.utils;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class RegexUtils {

    public static String regexWithRemove (String fullLine, String firstSearchRegex, String removeRegex) {
        String subLine;

        Pattern patternFirst = Pattern.compile(firstSearchRegex);
        Matcher matcherFirst = patternFirst.matcher(fullLine);
        if(matcherFirst.find()) {
            subLine = matcherFirst.group();
            Pattern patternForRemove = Pattern.compile(removeRegex);
            Matcher matcherForRemove = patternForRemove.matcher(subLine);
            matcherForRemove.find();
            String subResult = matcherForRemove.replaceAll("");
            String formattedResult = subResult.replaceAll("\\.",",")
                    .replaceAll("\\|","");

            return formattedResult;
        }

        return "";
    }

    public static String regex (String fullLine, String firstSearchRegex) {
        String subResult;

        Pattern patternFirst = Pattern.compile(firstSearchRegex);
        Matcher matcherFirst = patternFirst.matcher(fullLine);
        if(matcherFirst.find()) {
            subResult = matcherFirst.group();
            String formattedResult = subResult.replaceAll("\\.",",")
                    .replaceAll("\\|","");
            return formattedResult;
        }

        return "";
    }
}
