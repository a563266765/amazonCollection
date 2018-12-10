package com.biz;

import com.common.utils.DateUtil;
import com.dto.CommodityDTO;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;
import lombok.Getter;
import lombok.Setter;

import java.io.File;
import java.io.IOException;
import java.util.List;

class ExcelExportBiz {

    @Setter
    @Getter
    private CommodityDTO commodityDTO;
    public ExcelExportBiz(CommodityDTO commodityDTO) {
        this.commodityDTO = commodityDTO;
    }

    public void exportExcel(CommodityDTO commodityDTO, String filePath){

            try {
                List<String> bulletPoint;
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

}
