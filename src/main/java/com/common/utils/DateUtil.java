package com.common.utils;

import java.text.SimpleDateFormat;
import java.util.Calendar;

public class DateUtil {


    public String ymdFormat(){
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyyMMdd");
        return simpleDateFormat.format(System.currentTimeMillis());
    }

    public String hmsFormat(){
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("HHmmss");
        return simpleDateFormat.format(System.currentTimeMillis());
    }

    public String mdFormat(){
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("MMdd");
        return simpleDateFormat.format(System.currentTimeMillis());
    }

    public String yFormat(){
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy");
        return simpleDateFormat.format(System.currentTimeMillis());
    }

    public String y_M_dFormat(){
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
        return simpleDateFormat.format(System.currentTimeMillis());
    }

    public String dateYearsAlgorithm(int yearNum){

        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
        Calendar c = Calendar.getInstance();
        c.add(Calendar.YEAR,yearNum);
        return simpleDateFormat.format(c.getTime());

    }
}
