package com.biz;

import com.common.exception.ServiceException;
import com.common.exception.ServiceExceptionEnum;
import com.common.utils.DateUtil;
import com.dto.CommodityDTO;
import com.instruct.BaseInstruct;
import com.instruct.impl.AttributefieldImpl;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import org.jsoup.nodes.Document;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

@Service
public class CommodityInstruct implements BaseInstruct {

    /**
     * 爬取商品数据各字段信息功能实现类
     */
    @Autowired
    private AttributefieldImpl attributefield;

    public CommodityInstruct() {

    }

    @Override
    public void getPageSource(Document doc, String filePath, int lineSize, String path){

        CommodityDTO commodityDTO = new CommodityDTO();
        // 爬取商品页面各字段数据
        commodityDTO = this.getCrawlData(commodityDTO,doc,filePath,lineSize,path);
        WritableWorkbook book = null;
            try {
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
                book = Workbook.createWorkbook(new File(filePath + ".xls"));
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
                // 这里对颜色集合和尺寸集合的校验原因为亚马逊商品有可能出现无颜色选项或者无尺寸选项的情况，全部考虑到位包含上颜色尺寸都没有的情况
                if(commodityDTO.getColorList().size() !=0 && commodityDTO.getColorList() != null && commodityDTO.getSizeList().size()!= 0 && commodityDTO.getSizeList() !=null) {

                    // 以颜色和大小为维度创建所有同颜色不同大小和同大小不同颜色的子类
                    for (int i = 0; i < commodityDTO.getColorList().size(); i++) {

                        for (int j = 0; j < commodityDTO.getSizeList().size(); j++) {
                            // row为0是父类那行，需要单独与子类分开，父类子类分开的部分是不能公用的各类字段
                            if (row == 0) {
                                // 父类信息
                                this.parentInformation(row, ws);
                            } else {
                                // 子类信息
                                this.childInformation(row,ws,commodityDTO,i,j);
                            }
                            // 这里往后是可公用的各字段
                            // 根据采集到的商品描述校验，商品描述大于5条，只取后五条描述
                            this.getFiveBulletPoint(row,ws,commodityDTO);
                            this.getOtherReuseField(row, ws, commodityDTO);
                            row++;
                        }
                    }
                }else if(commodityDTO.getColorList().size() != 0 && commodityDTO.getColorList() != null){

                    // 上一个条件判断已经校验了颜色集合和尺寸集合，会进到这里说明了尺寸集合为null
                        for (int i = 0; i < commodityDTO.getColorList().size(); i++) {

                                // row为0是父类那行
                                if (row == 0) {
                                    // 父类信息
                                    this.parentInformation(row, ws);
                                } else {
                                    // 子类信息,这里给 j 传一个 0 进去在里面会有校验做不同的处理
                                    this.childInformation(row,ws,commodityDTO,i,0);
                                }
                            // 这里往后是可公用的各字段
                            // 根据采集到的商品描述校验，商品描述大于5条，只取后五条描述
                            this.getFiveBulletPoint(row,ws,commodityDTO);
                            this.getOtherReuseField(row, ws, commodityDTO);
                            row++;
                        }
                }else if(commodityDTO.getSizeList().size()!= 0 && commodityDTO.getSizeList() !=null){
                    // 上一个条件判断已经校验了颜色集合和尺寸集合，会进到这里说明了颜色集合为null
                        for (int j = 0; j < commodityDTO.getSizeList().size(); j++) {

                            // row为0是父类那行
                            if (row == 0) {
                                // 父类信息
                                this.parentInformation(row, ws);
                            } else {
                                // 子类信息,这里给 i 传一个 0 进去在里面会有校验做不同的处理
                                this.childInformation(row,ws,commodityDTO,0,j);
                            }
                            // 这里往后是可公用的各字段
                            // 根据采集到的商品描述校验，商品描述大于5条，只取后五条描述
                            this.getFiveBulletPoint(row,ws,commodityDTO);
                            this.getOtherReuseField(row, ws, commodityDTO);
                            row++;
                    }
                }else{
                    // 父类信息
                    this.parentInformation(row, ws);
                    // 这里往后是可公用的各字段
                    // 根据采集到的商品描述校验，商品描述大于5条，只取后五条描述
                    this.getFiveBulletPoint(row,ws,commodityDTO);
                    this.getOtherReuseField(row, ws, commodityDTO);
                }

            } catch (Exception e) {
                e.printStackTrace();
            }finally {
                try {
                    if(book != null){
                        book.write();
                        book.close();
                    }
                } catch (IOException | WriteException e) {
                    e.printStackTrace();
                }

            }

    }


    private void parentInformation(int row,WritableSheet ws) throws WriteException {

        Label item_sku = new Label(0, 3 + row, "ys" + DateUtil.ymdFormat() + "-1");
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
    }

    private void childInformation(int row,WritableSheet ws,CommodityDTO commodityDTO,int i, int j) throws WriteException {


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
        Label parent_sku = new Label(79, 3 + row, "ys" + DateUtil.ymdFormat() + "-1");
        Label relationship_type = new Label(80, 3 + row, "Variation");
        if(i == 0 && j != 0){
            Label color_name = new Label(95, 3 + row, "");
            Label color_map = new Label(96, 3 + row, "");
            Label item_sku = new Label(0, 3 + row, "ys" + DateUtil.ymdFormat() + "-" + commodityDTO.getLineSize() + "-"
                    + "-" + commodityDTO.getSizeList().get(j));
            Label size_name = new Label(120, 3 + row, commodityDTO.getSizeList().get(j));
            Label size_map = new Label(121, 3 + row, commodityDTO.getSizeList().get(j));
            ws.addCell(size_name);
            ws.addCell(size_map);
            ws.addCell(color_name);
            ws.addCell(color_map);
            ws.addCell(item_sku);
        }else if(j == 0 && i != 0) {
            Label item_sku = new Label(0, 3 + row, "ys" + DateUtil.ymdFormat() + "-" + commodityDTO.getLineSize() + "-"
                    + commodityDTO.getColorList().get(i));
            Label size_name = new Label(120, 3 + row, "");
            Label size_map = new Label(121, 3 + row, "");
            Label color_name = new Label(95, 3 + row, commodityDTO.getColorList().get(i));
            Label color_map = new Label(96, 3 + row, commodityDTO.getColorList().get(i));
            ws.addCell(color_name);
            ws.addCell(color_map);
            ws.addCell(size_name);
            ws.addCell(size_map);
            ws.addCell(item_sku);
        }else if( i == 0 && j == 0){
            Label item_sku = new Label(0, 3 + row, "ys" + DateUtil.ymdFormat() + "-" + commodityDTO.getLineSize());
            Label size_name = new Label(120, 3 + row, "");
            Label size_map = new Label(121, 3 + row, "");
            Label color_name = new Label(95, 3 + row, "");
            Label color_map = new Label(96, 3 + row, "");
            ws.addCell(color_name);
            ws.addCell(color_map);
            ws.addCell(size_name);
            ws.addCell(size_map);
            ws.addCell(item_sku);
        } else{
            Label item_sku = new Label(0, 3 + row, "ys" + DateUtil.ymdFormat() + "-" + commodityDTO.getLineSize() + "-"
                    + commodityDTO.getColorList().get(i) + "-" + commodityDTO.getSizeList().get(j));
            Label size_name = new Label(120, 3 + row, commodityDTO.getSizeList().get(j));
            Label size_map = new Label(121, 3 + row, commodityDTO.getSizeList().get(j));
            Label color_name = new Label(95, 3 + row, commodityDTO.getColorList().get(i));
            Label color_map = new Label(96, 3 + row, commodityDTO.getColorList().get(i));
            ws.addCell(color_name);
            ws.addCell(color_map);
            ws.addCell(size_name);
            ws.addCell(size_map);
            ws.addCell(item_sku);
        }
        ws.addCell(list_price);
        ws.addCell(standard_price);
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
    }

    private void getFiveBulletPoint(int row,WritableSheet ws,CommodityDTO commodityDTO) throws WriteException{

        int bulletPointSize = commodityDTO.getBulletPoint().size();
        List<String> bulletPoint = commodityDTO.getBulletPoint();
        if (bulletPointSize == 0){
            throw new ServiceException(ServiceExceptionEnum.BULLET_POINT_ERROR);
        }
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
        } else  {
            int b = 0;
            for (int a = 1; a <= bulletPointSize; a++) {
                Label bullet_point_not_null = new Label(50 + a, 3 + row, bulletPoint.get(b));
                ws.addCell(bullet_point_not_null);
                b++;
            }

        }

    }

    private void getOtherReuseField(int row,WritableSheet ws,CommodityDTO commodityDTO)throws WriteException{

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
    }

    private CommodityDTO getCrawlData(CommodityDTO commodityDTO,Document doc, String filePath, int lineSize, String path){

        Set<String> imgUrls = new HashSet<>();
        List<String> colorList = new ArrayList<>();
        List<String> sizeList = new ArrayList<>();
        List<String> bulletPoint = new ArrayList<>();
        StringBuilder itemType = new StringBuilder();
        // 尺寸集合
        commodityDTO.setSizeList(attributefield.sizeContainer(doc, sizeList));
        //imgUrl集合
        commodityDTO.setImgUrls(attributefield.imgUrl(doc, imgUrls, filePath, path));
        //金额获取
        commodityDTO.setSalePrice(attributefield.price(doc));
        //获取颜色
        commodityDTO.setColorList(attributefield.colorList(doc, colorList));
        List<String> itemAndGender = attributefield.productName(doc);
        commodityDTO.setItemName(itemAndGender.get(0));
        commodityDTO.setGender(itemAndGender.get(1).replace(" ", ""));
        commodityDTO.setSalePrice(attributefield.salePrice(doc));
        BigDecimal Magnification = new BigDecimal(1.4);
        BigDecimal standardPrice = new BigDecimal(commodityDTO.getSalePrice());
        standardPrice = standardPrice.multiply(Magnification).setScale(2, BigDecimal.ROUND_HALF_UP);
        commodityDTO.setStandardPrice(Double.toString(standardPrice.doubleValue()));
        commodityDTO.setListPrice(Double.toString(standardPrice.multiply(Magnification).setScale(2, BigDecimal.ROUND_HALF_UP).doubleValue()));
        commodityDTO.setBulletPoint(attributefield.bulletPointList(doc, bulletPoint));
        commodityDTO.setItemType(attributefield.itemType(doc, itemType));
        commodityDTO.setLineSize(lineSize);
        return commodityDTO;
    }

}
