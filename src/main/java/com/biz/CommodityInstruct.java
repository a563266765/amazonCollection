package com.biz;

import com.common.utils.DateUtil;
import com.dto.CommodityDTO;
import com.instruct.BaseInstruct;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;
import org.apache.commons.lang.StringUtils;
import org.jsoup.nodes.Document;
import org.jsoup.select.Elements;
import org.springframework.stereotype.Service;

import java.io.*;
import java.math.BigDecimal;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLEncoder;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@Service
public class CommodityInstruct implements BaseInstruct {

    public CommodityInstruct() {

    }

    @Override
    public void getPageSource(Document doc, String filePath, int lineSize, String path){

        CommodityDTO commodityDTO = new CommodityDTO();
        Set<String> imgUrls = new HashSet<>();
        List<String> colorList = new ArrayList<>();
        List<String> sizeList = new ArrayList<>();
        List<String> bulletPoint = new ArrayList<>();
        StringBuilder itemType = new StringBuilder();
        // 尺寸集合
        commodityDTO.setSizeList(sizeContainer(doc, sizeList));
        //imgUrl集合
        commodityDTO.setImgUrls(imgUrl(doc, imgUrls, filePath, path));
        //金额获取
        commodityDTO.setSalePrice(price(doc));
        //获取颜色
        commodityDTO.setColorList(colorList(doc, colorList));
        List<String> itemAndGender = productName(doc);
        commodityDTO.setItemName(itemAndGender.get(0));
        commodityDTO.setGender(itemAndGender.get(1).replace(" ", ""));
        commodityDTO.setSalePrice(salePrice(doc));
        BigDecimal Magnification = new BigDecimal(1.4);
        BigDecimal standardPrice = new BigDecimal(commodityDTO.getSalePrice());
        standardPrice = standardPrice.multiply(Magnification).setScale(2, BigDecimal.ROUND_HALF_UP);
        commodityDTO.setStandardPrice(Double.toString(standardPrice.doubleValue()));
        commodityDTO.setListPrice(Double.toString(standardPrice.multiply(Magnification).setScale(2, BigDecimal.ROUND_HALF_UP).doubleValue()));
        commodityDTO.setBulletPoint(bulletPointList(doc, bulletPoint));
        commodityDTO.setItemType(itemType(doc, itemType));
        commodityDTO.setLineSize(lineSize);
            try {
                DateUtil dateUtil = new DateUtil();
                File testExcel = new File("C:\\Users\\Administrator\\Desktop\\amazon\\ys0727.xls");
                //读取ys0727.xls
                Workbook  wb = Workbook.getWorkbook(testExcel);
                // 获取ys0727.xls第一个sheet
                Sheet read= wb.getSheet(0);
                // 在桌面新建一个ys4位月日的excel
                int lastIndex = filePath.lastIndexOf("\\");
                String pathName = filePath.substring(0,lastIndex);
                File excel = new File(pathName + "\\");
                if(!excel.exists()){
                    excel.mkdirs();
                }
                WritableWorkbook book = Workbook.createWorkbook(new File(filePath + ".xls"));
                WritableSheet ws = book.createSheet("sheet1",0);
                // 标头前三行
                for(int  i = 0; i < 3 ; i++) {
                    Cell columNum[] = read.getRow(i);
                    for (int  j = 0; j < columNum.length; j++) {
                        Label label = new Label( j, i, columNum[j].getContents());
                        ws.addCell(label);
                    }
                }
                wb.close();

                int row = 0;
                if(commodityDTO.getColorList().size() !=0 && commodityDTO.getColorList() != null && commodityDTO.getSizeList().size()!= 0 && commodityDTO.getSizeList() !=null) {

                    for (int i = 0; i < commodityDTO.getColorList().size(); i++) {

                        for (int j = 0; j < commodityDTO.getSizeList().size(); j++) {

                            // row为0是父类那行
                            if (row == 0) {
                                Label item_sku = new Label(0, 3 + row, "ys" + dateUtil.ymdFormat() + "-1");
                                Label external_product_id_type = new Label(3, 3 + row, "");
                                Label fulfillment_latency = new Label(12, 3 + row, "");
                                Label quantity = new Label(17, 3 + row, "");
                                Label sale_price = new Label(16, 3 + row, "");
                                Label sale_from_date = new Label(18, 3 + row, "");
                                Label sale_end_date = new Label(19, 3 + row, "");
                                Label merchant_shipping_group_name = new Label(27, 3 + row, "");
                                Label parent_child = new Label(78, 3 + row, "parent");
                                Label parent_sku = new Label(79, 3 + row, "");
                                Label relationship_type = new Label(80, 3 + row, "");
                                Label size_name = new Label(120, 3 + row, "");
                                Label size_map = new Label(121, 3 + row, "");
                                Label color_name = new Label(95, 3 + row, "");
                                Label color_map = new Label(96, 3 + row, "");
                                ws.addCell(color_name);
                                ws.addCell(color_map);
                                ws.addCell(size_name);
                                ws.addCell(size_map);
                                ws.addCell(relationship_type);
                                ws.addCell(parent_sku);
                                ws.addCell(parent_child);
                                ws.addCell(merchant_shipping_group_name);
                                ws.addCell(sale_from_date);
                                ws.addCell(sale_end_date);
                                ws.addCell(sale_price);
                                ws.addCell(quantity);
                                ws.addCell(fulfillment_latency);
                                ws.addCell(item_sku);
                                ws.addCell(external_product_id_type);
                            } else {
                                Label item_sku = new Label(0, 3 + row, "ys" + dateUtil.ymdFormat() + "-" + commodityDTO.getLineSize() + "-"
                                        + commodityDTO.getColorList().get(i) + "-" + commodityDTO.getSizeList().get(j));
                                Label external_product_id_type = new Label(3, 3 + row, "UPC");
                                Label fulfillment_latency = new Label(12, 3 + row, "8");
                                Label quantity = new Label(16, 3 + row, "100");
                                Label sale_price = new Label(17, 3 + row, commodityDTO.getSalePrice());
                                Label standard_price = new Label(9, 3 + row, commodityDTO.getStandardPrice());
                                Label list_price = new Label(10, 3 + row, commodityDTO.getListPrice());
                                Label sale_from_date = new Label(18, 3 + row, commodityDTO.getSaleFromDate());
                                Label sale_end_date = new Label(19, 3 + row, commodityDTO.getSaleEndDate());
                                Label merchant_shipping_group_name = new Label(27, 3 + row, commodityDTO.getShippingTemplate());
                                Label parent_child = new Label(78, 3 + row, "child");
                                Label parent_sku = new Label(79, 3 + row, "ys" + dateUtil.ymdFormat() + "-1");
                                Label relationship_type = new Label(80, 3 + row, "Variation");
                                Label color_name = new Label(95, 3 + row, commodityDTO.getColorList().get(i));
                                Label color_map = new Label(96, 3 + row, commodityDTO.getColorList().get(i));
                                Label size_name = new Label(120, 3 + row, commodityDTO.getSizeList().get(j));
                                Label size_map = new Label(121, 3 + row, commodityDTO.getSizeList().get(j));

                                ws.addCell(list_price);
                                ws.addCell(standard_price);
                                ws.addCell(size_name);
                                ws.addCell(size_map);
                                ws.addCell(color_name);
                                ws.addCell(color_map);
                                ws.addCell(relationship_type);
                                ws.addCell(parent_sku);
                                ws.addCell(parent_child);
                                ws.addCell(merchant_shipping_group_name);
                                ws.addCell(sale_from_date);
                                ws.addCell(sale_end_date);
                                ws.addCell(sale_price);
                                ws.addCell(quantity);
                                ws.addCell(fulfillment_latency);
                                ws.addCell(external_product_id_type);
                                ws.addCell(item_sku);
                            }
                            int bulletPointSize = commodityDTO.getBulletPoint().size();
                            bulletPoint = commodityDTO.getBulletPoint();
                            if (commodityDTO.getBulletPoint().size() >= 5) {
                                Label bullet_point1 = new Label(51, 3 + row, bulletPoint.get(bulletPointSize - 5));
                                Label bullet_point2 = new Label(52, 3 + row, bulletPoint.get(bulletPointSize - 4));
                                Label bullet_point3 = new Label(53, 3 + row, bulletPoint.get(bulletPointSize - 3));
                                Label bullet_point4 = new Label(54, 3 + row, bulletPoint.get(bulletPointSize - 2));
                                Label bullet_point5 = new Label(55, 3 + row, bulletPoint.get(bulletPointSize - 1));
                                ws.addCell(bullet_point1);
                                ws.addCell(bullet_point2);
                                ws.addCell(bullet_point3);
                                ws.addCell(bullet_point4);
                                ws.addCell(bullet_point5);
                            } else if (bulletPointSize != 0) {
                                int b = 0;
                                for (int a = 1; a <= bulletPointSize; a++) {
                                    Label bullet_point_not_null = new Label(50 + a, 3 + row, bulletPoint.get(b));
                                    ws.addCell(bullet_point_not_null);
                                    b++;
                                }

                            }

                            Label item_name = new Label(1, 3 + row, commodityDTO.getItemName());
                            Label brand_name = new Label(4, 3 + row, "Surprise4U");
                            Label variation_theme = new Label(81, 3 + row, "SizeColor");
                            Label item_type = new Label(6, 3 + row, commodityDTO.getItemType());
                            Label external_product_id = new Label(2, 3 + row, commodityDTO.getGender());
                            Label item_package_quantity = new Label(11, 3 + row, commodityDTO.getGender());
                            ws.addCell(item_package_quantity);
                            ws.addCell(external_product_id);
                            ws.addCell(item_type);
                            ws.addCell(variation_theme);
                            ws.addCell(item_name);
                            ws.addCell(brand_name);

                            row++;
                        }
                    }
                }else if(commodityDTO.getColorList().size() != 0 && commodityDTO.getColorList() != null){
                    {
                        for (int i = 0; i < commodityDTO.getColorList().size(); i++) {

                                // row为0是父类那行
                                if (row == 0) {
                                    Label item_sku = new Label(0, 3 + row, "ys" + dateUtil.ymdFormat() + "-1");
                                    Label external_product_id_type = new Label(3, 3 + row, "");
                                    Label fulfillment_latency = new Label(12, 3 + row, "");
                                    Label quantity = new Label(17, 3 + row, "");
                                    Label sale_price = new Label(16, 3 + row, "");
                                    Label sale_from_date = new Label(18, 3 + row, "");
                                    Label sale_end_date = new Label(19, 3 + row, "");
                                    Label merchant_shipping_group_name = new Label(27, 3 + row, "");
                                    Label parent_child = new Label(78, 3 + row, "parent");
                                    Label parent_sku = new Label(79, 3 + row, "");
                                    Label relationship_type = new Label(80, 3 + row, "");
                                    Label size_name = new Label(120, 3 + row, "");
                                    Label size_map = new Label(121, 3 + row, "");
                                    Label color_name = new Label(95, 3 + row, "");
                                    Label color_map = new Label(96, 3 + row, "");
                                    ws.addCell(color_name);
                                    ws.addCell(color_map);
                                    ws.addCell(size_name);
                                    ws.addCell(size_map);
                                    ws.addCell(relationship_type);
                                    ws.addCell(parent_sku);
                                    ws.addCell(parent_child);
                                    ws.addCell(merchant_shipping_group_name);
                                    ws.addCell(sale_from_date);
                                    ws.addCell(sale_end_date);
                                    ws.addCell(sale_price);
                                    ws.addCell(quantity);
                                    ws.addCell(fulfillment_latency);
                                    ws.addCell(item_sku);
                                    ws.addCell(external_product_id_type);
                                } else {
                                    Label item_sku = new Label(0, 3 + row, "ys" + dateUtil.ymdFormat() + "-" + commodityDTO.getLineSize() + "-"
                                            + commodityDTO.getColorList().get(i));
                                    Label external_product_id_type = new Label(3, 3 + row, "UPC");
                                    Label fulfillment_latency = new Label(12, 3 + row, "8");
                                    Label quantity = new Label(16, 3 + row, "100");
                                    Label sale_price = new Label(17, 3 + row, commodityDTO.getSalePrice());
                                    Label standard_price = new Label(9, 3 + row, commodityDTO.getStandardPrice());
                                    Label list_price = new Label(10, 3 + row, commodityDTO.getListPrice());
                                    Label sale_from_date = new Label(18, 3 + row, commodityDTO.getSaleFromDate());
                                    Label sale_end_date = new Label(19, 3 + row, commodityDTO.getSaleEndDate());
                                    Label merchant_shipping_group_name = new Label(27, 3 + row, commodityDTO.getShippingTemplate());
                                    Label parent_child = new Label(78, 3 + row, "child");
                                    Label parent_sku = new Label(79, 3 + row, "ys" + dateUtil.ymdFormat() + "-1");
                                    Label relationship_type = new Label(80, 3 + row, "Variation");
                                    Label color_name = new Label(95, 3 + row, commodityDTO.getColorList().get(i));
                                    Label color_map = new Label(96, 3 + row, commodityDTO.getColorList().get(i));
                                    Label size_name = new Label(120, 3 + row, "");
                                    Label size_map = new Label(121, 3 + row, "");

                                    ws.addCell(list_price);
                                    ws.addCell(standard_price);
                                    ws.addCell(size_name);
                                    ws.addCell(size_map);
                                    ws.addCell(color_name);
                                    ws.addCell(color_map);
                                    ws.addCell(relationship_type);
                                    ws.addCell(parent_sku);
                                    ws.addCell(parent_child);
                                    ws.addCell(merchant_shipping_group_name);
                                    ws.addCell(sale_from_date);
                                    ws.addCell(sale_end_date);
                                    ws.addCell(sale_price);
                                    ws.addCell(quantity);
                                    ws.addCell(fulfillment_latency);
                                    ws.addCell(external_product_id_type);
                                    ws.addCell(item_sku);
                                }
                                int bulletPointSize = commodityDTO.getBulletPoint().size();
                                bulletPoint = commodityDTO.getBulletPoint();
                                if (commodityDTO.getBulletPoint().size() >= 5) {
                                    Label bullet_point1 = new Label(51, 3 + row, bulletPoint.get(bulletPointSize - 5));
                                    Label bullet_point2 = new Label(52, 3 + row, bulletPoint.get(bulletPointSize - 4));
                                    Label bullet_point3 = new Label(53, 3 + row, bulletPoint.get(bulletPointSize - 3));
                                    Label bullet_point4 = new Label(54, 3 + row, bulletPoint.get(bulletPointSize - 2));
                                    Label bullet_point5 = new Label(55, 3 + row, bulletPoint.get(bulletPointSize - 1));
                                    ws.addCell(bullet_point1);
                                    ws.addCell(bullet_point2);
                                    ws.addCell(bullet_point3);
                                    ws.addCell(bullet_point4);
                                    ws.addCell(bullet_point5);
                                } else if (bulletPointSize != 0) {
                                    int b = 0;
                                    for (int a = 1; a <= bulletPointSize; a++) {
                                        Label bullet_point_not_null = new Label(50 + a, 3 + row, bulletPoint.get(b));
                                        ws.addCell(bullet_point_not_null);
                                        b++;
                                    }

                                }

                                Label item_name = new Label(1, 3 + row, commodityDTO.getItemName());
                                Label brand_name = new Label(4, 3 + row, "Surprise4U");
                                Label variation_theme = new Label(81, 3 + row, "SizeColor");
                                Label item_type = new Label(6, 3 + row, commodityDTO.getItemType());
                                Label external_product_id = new Label(2, 3 + row, commodityDTO.getGender());
                                Label item_package_quantity = new Label(11, 3 + row, commodityDTO.getGender());
                                ws.addCell(item_package_quantity);
                                ws.addCell(external_product_id);
                                ws.addCell(item_type);
                                ws.addCell(variation_theme);
                                ws.addCell(item_name);
                                ws.addCell(brand_name);

                                row++;

                        }
                    }

                }else if(commodityDTO.getSizeList().size()!= 0 && commodityDTO.getSizeList() !=null){

                        for (int j = 0; j < commodityDTO.getSizeList().size(); j++) {

                            // row为0是父类那行
                            if (row == 0) {
                                Label item_sku = new Label(0, 3 + row, "ys" + dateUtil.ymdFormat() + "-1");
                                Label external_product_id_type = new Label(3, 3 + row, "");
                                Label fulfillment_latency = new Label(12, 3 + row, "");
                                Label quantity = new Label(17, 3 + row, "");
                                Label sale_price = new Label(16, 3 + row, "");
                                Label sale_from_date = new Label(18, 3 + row, "");
                                Label sale_end_date = new Label(19, 3 + row, "");
                                Label merchant_shipping_group_name = new Label(27, 3 + row, "");
                                Label parent_child = new Label(78, 3 + row, "parent");
                                Label parent_sku = new Label(79, 3 + row, "");
                                Label relationship_type = new Label(80, 3 + row, "");
                                Label size_name = new Label(120, 3 + row, "");
                                Label size_map = new Label(121, 3 + row, "");
                                Label color_name = new Label(95, 3 + row, "");
                                Label color_map = new Label(96, 3 + row, "");
                                ws.addCell(color_name);
                                ws.addCell(color_map);
                                ws.addCell(size_name);
                                ws.addCell(size_map);
                                ws.addCell(relationship_type);
                                ws.addCell(parent_sku);
                                ws.addCell(parent_child);
                                ws.addCell(merchant_shipping_group_name);
                                ws.addCell(sale_from_date);
                                ws.addCell(sale_end_date);
                                ws.addCell(sale_price);
                                ws.addCell(quantity);
                                ws.addCell(fulfillment_latency);
                                ws.addCell(item_sku);
                                ws.addCell(external_product_id_type);
                            } else {
                                Label item_sku = new Label(0, 3 + row, "ys" + dateUtil.ymdFormat() + "-" + commodityDTO.getLineSize() + "-"
                                         + "-" + commodityDTO.getSizeList().get(j));
                                Label external_product_id_type = new Label(3, 3 + row, "UPC");
                                Label fulfillment_latency = new Label(12, 3 + row, "8");
                                Label quantity = new Label(16, 3 + row, "100");
                                Label sale_price = new Label(17, 3 + row, commodityDTO.getSalePrice());
                                Label standard_price = new Label(9, 3 + row, commodityDTO.getStandardPrice());
                                Label list_price = new Label(10, 3 + row, commodityDTO.getListPrice());
                                Label sale_from_date = new Label(18, 3 + row, commodityDTO.getSaleFromDate());
                                Label sale_end_date = new Label(19, 3 + row, commodityDTO.getSaleEndDate());
                                Label merchant_shipping_group_name = new Label(27, 3 + row, commodityDTO.getShippingTemplate());
                                Label parent_child = new Label(78, 3 + row, "child");
                                Label parent_sku = new Label(79, 3 + row, "ys" + dateUtil.ymdFormat() + "-1");
                                Label relationship_type = new Label(80, 3 + row, "Variation");
                                Label color_name = new Label(95, 3 + row, "");
                                Label color_map = new Label(96, 3 + row, "");
                                Label size_name = new Label(120, 3 + row, commodityDTO.getSizeList().get(j));
                                Label size_map = new Label(121, 3 + row, commodityDTO.getSizeList().get(j));

                                ws.addCell(list_price);
                                ws.addCell(standard_price);
                                ws.addCell(size_name);
                                ws.addCell(size_map);
                                ws.addCell(color_name);
                                ws.addCell(color_map);
                                ws.addCell(relationship_type);
                                ws.addCell(parent_sku);
                                ws.addCell(parent_child);
                                ws.addCell(merchant_shipping_group_name);
                                ws.addCell(sale_from_date);
                                ws.addCell(sale_end_date);
                                ws.addCell(sale_price);
                                ws.addCell(quantity);
                                ws.addCell(fulfillment_latency);
                                ws.addCell(external_product_id_type);
                                ws.addCell(item_sku);
                            }
                            int bulletPointSize = commodityDTO.getBulletPoint().size();
                            bulletPoint = commodityDTO.getBulletPoint();
                            if (commodityDTO.getBulletPoint().size() >= 5) {
                                Label bullet_point1 = new Label(51, 3 + row, bulletPoint.get(bulletPointSize - 5));
                                Label bullet_point2 = new Label(52, 3 + row, bulletPoint.get(bulletPointSize - 4));
                                Label bullet_point3 = new Label(53, 3 + row, bulletPoint.get(bulletPointSize - 3));
                                Label bullet_point4 = new Label(54, 3 + row, bulletPoint.get(bulletPointSize - 2));
                                Label bullet_point5 = new Label(55, 3 + row, bulletPoint.get(bulletPointSize - 1));
                                ws.addCell(bullet_point1);
                                ws.addCell(bullet_point2);
                                ws.addCell(bullet_point3);
                                ws.addCell(bullet_point4);
                                ws.addCell(bullet_point5);
                            } else if (bulletPointSize != 0) {
                                int b = 0;
                                for (int a = 1; a <= bulletPointSize; a++) {
                                    Label bullet_point_not_null = new Label(50 + a, 3 + row, bulletPoint.get(b));
                                    ws.addCell(bullet_point_not_null);
                                    b++;
                                }

                            }

                            Label item_name = new Label(1, 3 + row, commodityDTO.getItemName());
                            Label brand_name = new Label(4, 3 + row, "Surprise4U");
                            Label variation_theme = new Label(81, 3 + row, "SizeColor");
                            Label item_type = new Label(6, 3 + row, commodityDTO.getItemType());
                            Label external_product_id = new Label(2, 3 + row, commodityDTO.getGender());
                            Label item_package_quantity = new Label(11, 3 + row, commodityDTO.getGender());
                            ws.addCell(item_package_quantity);
                            ws.addCell(external_product_id);
                            ws.addCell(item_type);
                            ws.addCell(variation_theme);
                            ws.addCell(item_name);
                            ws.addCell(brand_name);

                            row++;

                    }
                }else{

                            // row为0是父类那行
                            if (row == 0) {
                                Label item_sku = new Label(0, 3 + row, "ys" + dateUtil.ymdFormat() + "-1");
                                Label external_product_id_type = new Label(3, 3 + row, "");
                                Label fulfillment_latency = new Label(12, 3 + row, "");
                                Label quantity = new Label(17, 3 + row, "");
                                Label sale_price = new Label(16, 3 + row, "");
                                Label sale_from_date = new Label(18, 3 + row, "");
                                Label sale_end_date = new Label(19, 3 + row, "");
                                Label merchant_shipping_group_name = new Label(27, 3 + row, "");
                                Label parent_child = new Label(78, 3 + row, "parent");
                                Label parent_sku = new Label(79, 3 + row, "");
                                Label relationship_type = new Label(80, 3 + row, "");
                                Label size_name = new Label(120, 3 + row, "");
                                Label size_map = new Label(121, 3 + row, "");
                                Label color_name = new Label(95, 3 + row, "");
                                Label color_map = new Label(96, 3 + row, "");
                                ws.addCell(color_name);
                                ws.addCell(color_map);
                                ws.addCell(size_name);
                                ws.addCell(size_map);
                                ws.addCell(relationship_type);
                                ws.addCell(parent_sku);
                                ws.addCell(parent_child);
                                ws.addCell(merchant_shipping_group_name);
                                ws.addCell(sale_from_date);
                                ws.addCell(sale_end_date);
                                ws.addCell(sale_price);
                                ws.addCell(quantity);
                                ws.addCell(fulfillment_latency);
                                ws.addCell(item_sku);
                                ws.addCell(external_product_id_type);
                            } else {
                                Label item_sku = new Label(0, 3 + row, "ys" + dateUtil.ymdFormat() + "-" + commodityDTO.getLineSize());
                                Label external_product_id_type = new Label(3, 3 + row, "UPC");
                                Label fulfillment_latency = new Label(12, 3 + row, "8");
                                Label quantity = new Label(16, 3 + row, "100");
                                Label sale_price = new Label(17, 3 + row, commodityDTO.getSalePrice());
                                Label standard_price = new Label(9, 3 + row, commodityDTO.getStandardPrice());
                                Label list_price = new Label(10, 3 + row, commodityDTO.getListPrice());
                                Label sale_from_date = new Label(18, 3 + row, commodityDTO.getSaleFromDate());
                                Label sale_end_date = new Label(19, 3 + row, commodityDTO.getSaleEndDate());
                                Label merchant_shipping_group_name = new Label(27, 3 + row, commodityDTO.getShippingTemplate());
                                Label parent_child = new Label(78, 3 + row, "child");
                                Label parent_sku = new Label(79, 3 + row, "ys" + dateUtil.ymdFormat() + "-1");
                                Label relationship_type = new Label(80, 3 + row, "Variation");
                                Label color_name = new Label(95, 3 + row, "");
                                Label color_map = new Label(96, 3 + row, "");
                                Label size_name = new Label(120, 3 + row, "");
                                Label size_map = new Label(121, 3 + row, "");

                                ws.addCell(list_price);
                                ws.addCell(standard_price);
                                ws.addCell(size_name);
                                ws.addCell(size_map);
                                ws.addCell(color_name);
                                ws.addCell(color_map);
                                ws.addCell(relationship_type);
                                ws.addCell(parent_sku);
                                ws.addCell(parent_child);
                                ws.addCell(merchant_shipping_group_name);
                                ws.addCell(sale_from_date);
                                ws.addCell(sale_end_date);
                                ws.addCell(sale_price);
                                ws.addCell(quantity);
                                ws.addCell(fulfillment_latency);
                                ws.addCell(external_product_id_type);
                                ws.addCell(item_sku);
                            }
                            int bulletPointSize = commodityDTO.getBulletPoint().size();
                            bulletPoint = commodityDTO.getBulletPoint();
                            if (commodityDTO.getBulletPoint().size() >= 5) {
                                Label bullet_point1 = new Label(51, 3 + row, bulletPoint.get(bulletPointSize - 5));
                                Label bullet_point2 = new Label(52, 3 + row, bulletPoint.get(bulletPointSize - 4));
                                Label bullet_point3 = new Label(53, 3 + row, bulletPoint.get(bulletPointSize - 3));
                                Label bullet_point4 = new Label(54, 3 + row, bulletPoint.get(bulletPointSize - 2));
                                Label bullet_point5 = new Label(55, 3 + row, bulletPoint.get(bulletPointSize - 1));
                                ws.addCell(bullet_point1);
                                ws.addCell(bullet_point2);
                                ws.addCell(bullet_point3);
                                ws.addCell(bullet_point4);
                                ws.addCell(bullet_point5);
                            } else if (bulletPointSize != 0) {
                                int b = 0;
                                for (int a = 1; a <= bulletPointSize; a++) {
                                    Label bullet_point_not_null = new Label(50 + a, 3 + row, bulletPoint.get(b));
                                    ws.addCell(bullet_point_not_null);
                                    b++;
                                }

                            }

                            Label item_name = new Label(1, 3 + row, commodityDTO.getItemName());
                            Label brand_name = new Label(4, 3 + row, "Surprise4U");
                            Label variation_theme = new Label(81, 3 + row, "SizeColor");
                            Label item_type = new Label(6, 3 + row, commodityDTO.getItemType());
                            Label external_product_id = new Label(2, 3 + row, commodityDTO.getGender());
                            Label item_package_quantity = new Label(11, 3 + row, commodityDTO.getGender());
                            ws.addCell(item_package_quantity);
                            ws.addCell(external_product_id);
                            ws.addCell(item_type);
                            ws.addCell(variation_theme);
                            ws.addCell(item_name);
                            ws.addCell(brand_name);

                            row++;
                }
                book.write();
                book.close();
            } catch (IOException e) {
                e.printStackTrace();
            } catch (RowsExceededException e) {
                e.printStackTrace();
            } catch (WriteException e) {
                e.printStackTrace();
            } catch (Exception e) {
                e.printStackTrace();
            }

    }

    private Set<String> imgUrl (Document doc, Set < String > imgUrls, String txtFilePath, String path){

        imgUrls = imgSize1000(doc, imgUrls);
        Iterator showIterator = imgUrls.iterator();
        if (StringUtils.isNotBlank(String.valueOf(imgUrls))) {
            String imgUrlPath = txtFilePath + "-img" + "\\";
            File files = new File(imgUrlPath);
            if (!files.exists()) {
                files.mkdirs();
            }
            int imgSize = 1;
            while (showIterator.hasNext()) {
                try {
                    imgDownload(imgUrlPath, showIterator.next().toString(), imgSize, path);
                    imgSize++;
                } catch (UnsupportedEncodingException e) {
                    System.out.println("图片下载异常！！！");
                    e.printStackTrace();
                }
            }
        }
        return imgUrls;
    }

    private String price (Document doc){
        return doc.select("span#priceblock_ourprice").text();
    }

    private static List<String> colorList (Document doc, List < String > colorList){

        Elements colors = doc.select("img.imgSwatch");

        for (int i = 0; i < colors.size(); i++) {

            //颜色代码截取
            String color = colors.get(i).attr("alt");
            colorList.add(color);
        }
        return colorList;
    }

    private List<String> sizeContainer (Document doc, List < String > sizeList){

        Elements sizeElements = doc.select("select.a-native-dropdown > option");

        for (int i = 0; i < sizeElements.size(); i++) {

            //尺寸代码截取
            String size = sizeElements.get(i).attr("data-a-html-content");
            if (size != null && !"".equals(size)) {
                sizeList.add(size);
            }
        }
        return sizeList;
    }

    private List<String> productName (Document doc){
        List<String> itemAndGender = new ArrayList<String>(2);
        Elements sizeElements = doc.select("SPAN#productTitle");

        String item_name = sizeElements.text();
        String[] itemList = item_name.split(" ");
        StringBuilder item = new StringBuilder();
        for (int i = 0; i < itemList.length; i++) {
            if (i == 0) {
                item.append("SP4U");
            } else {
                item.append(itemList[i]);
            }
            item.append(" ");
        }
        itemAndGender.add(item.toString());
        itemAndGender.add(itemList[1]);
        return itemAndGender;
    }

    private String salePrice (Document doc){

        Elements sizeElements = doc.select("SPAN#priceblock_ourprice");

        String item_name = sizeElements.text().replace("$", "");
        String[] itemName = item_name.split("-");
        return itemName[0].replace(" ", "");
    }

    private List<String> bulletPointList (Document doc, List < String > bulletPoint){

        Elements sizeElements = doc.select("div#feature-bullets > ul > li");
        for (int i = 0; i < sizeElements.size(); i++) {

            //卖点代码截取
            String size = sizeElements.get(i).text();
            if (size != null && !"".equals(size)) {
                bulletPoint.add(size);
            }
        }
        return bulletPoint;
    }

    private String itemType (Document doc, StringBuilder itemType ){

        Elements itemTypeElements = doc.select("SPAN.a-list-item > a");
        for (int i = 0; i < itemTypeElements.size(); i++) {

            String item_name = itemTypeElements.get(i).text();
            if (!"".equals(item_name) && item_name != null) {
                itemType.append(item_name.replaceAll(" ", ""));
                if (i != (itemTypeElements.size() - 1) && !"".equals(itemTypeElements.get(i + 1).text()) && itemTypeElements.get(i + 1).text() != null) {
                    itemType.append("-");
                }
            }
        }
        return itemType.toString();
    }

    private void imgDownload (String filePath, String imgUrl,int imgSize, String path)throws
            UnsupportedEncodingException {


        //图片url中的前面部分：例如"http://images.csdn.net/"
        String beforeUrl = imgUrl.substring(0, imgUrl.lastIndexOf("/") + 1);
        //图片url中的后面部分：例如“20150529/PP6A7429_副本1.jpg”
        String fileName = imgUrl.substring(imgUrl.lastIndexOf("/") + 1);
        //编码之后的fileName，空格会变成字符"+"
        String newFileName = URLEncoder.encode(fileName, "UTF-8");
        //把编码之后的fileName中的字符"+"，替换为UTF-8中的空格表示："%20"
        newFileName = newFileName.replaceAll("\\+", "\\%20");
        //编码之后的url
        imgUrl = beforeUrl + newFileName;
        String suffix = imgUrl.substring(imgUrl.lastIndexOf("."));
        try {
            //创建文件目录
            File files = new File(filePath);
            if (!files.exists()) {
                files.mkdirs();
            }
            //获取下载地址
            URL url = new URL(imgUrl);
            //链接网络地址
            HttpURLConnection connection = (HttpURLConnection) url.openConnection();
            //获取链接的输出流
            InputStream is = connection.getInputStream();
            //创建文件，fileName为编码之前的文件名
            File file = new File(filePath + path + "-" + imgSize + suffix);
            //根据输入流写入文件
            FileOutputStream out = new FileOutputStream(file);
            int i = 0;
            while ((i = is.read()) != -1) {
                out.write(i);
            }
            out.close();
            is.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

    }
    private Set<String> imgSize1000 (Document doc, Set < String > imgUrls ){
        Pattern patten = Pattern.compile("https://([\\w-]+\\.)+[\\w-]+(/[\\w- ./?%&=#]*)1\\d+_.jpg");
        Matcher m = patten.matcher(doc.toString());
        for (int i = 0; i < 100; i++) {

            if (m.find()) {
                imgUrls.add(m.group());
            } else {
                break;
            }
        }
        return imgUrls;
    }
}
