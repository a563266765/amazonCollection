package com.instruct.impl;

import com.instruct.Attributefield;
import org.apache.commons.lang.StringUtils;
import org.jsoup.nodes.Document;
import org.jsoup.select.Elements;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.UnsupportedEncodingException;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 爬取商品数据各字段信息功能实现类
 * projectName : demo
 *
 * @author yangshuai
 * @version ID : V1.0 , 2018/12/1718:48
 */
public class AttributefieldImpl implements Attributefield {

    @Override
    public Set<String> imgUrl (Document doc, Set < String > imgUrls, String txtFilePath, String path){

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

    @Override
    public String price (Document doc){
        return doc.select("span#priceblock_ourprice").text();
    }

    @Override
    public List<String> colorList (Document doc, List < String > colorList){

        Elements colors = doc.select("img.imgSwatch");

        for (int i = 0; i < colors.size(); i++) {

            //颜色代码截取
            String color = colors.get(i).attr("alt");
            colorList.add(color);
        }
        return colorList;
    }

    @Override
    public List<String> sizeContainer (Document doc, List < String > sizeList){

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

    @Override
    public List<String> productName (Document doc){
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

    @Override
    public String salePrice (Document doc){

        Elements sizeElements = doc.select("SPAN#priceblock_ourprice");

        String item_name = sizeElements.text().replace("$", "");
        String[] itemName = item_name.split("-");
        return itemName[0].replace(" ", "");
    }

    @Override
    public List<String> bulletPointList (Document doc, List < String > bulletPoint){

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

    @Override
    public String itemType (Document doc, StringBuilder itemType ){

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
