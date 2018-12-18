package com.common.utils;

import java.text.SimpleDateFormat;
import java.util.Calendar;

public class DateUtil {


    public static String ymdFormat(){
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyyMMdd");
        return simpleDateFormat.format(System.currentTimeMillis());
    }

    public static String hmsFormat(){
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("HHmmss");
        return simpleDateFormat.format(System.currentTimeMillis());
    }

    public static String mdFormat(){
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("MMdd");
        return simpleDateFormat.format(System.currentTimeMillis());
    }

    public static String yFormat(){
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy");
        return simpleDateFormat.format(System.currentTimeMillis());
    }

    public static String y_M_dFormat(){
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
        return simpleDateFormat.format(System.currentTimeMillis());
    }

    public static String dateYearsAlgorithm(int yearNum){

        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
        Calendar c = Calendar.getInstance();
        c.add(Calendar.YEAR,yearNum);
        return simpleDateFormat.format(c.getTime());

    }
}
