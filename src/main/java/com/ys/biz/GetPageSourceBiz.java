package com.ys.biz;

import com.ys.common.exception.ServiceException;
import com.ys.common.exception.ServiceExceptionEnum;
import com.ys.common.utils.DateUtil;
import com.ys.control.BaseInstructControl;
import com.ys.instruct.BaseInstruct;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang.StringUtils;
import org.jsoup.Connection;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import javax.annotation.PostConstruct;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStreamReader;

@Service
@Slf4j
public class GetPageSourceBiz {

    @Autowired
    private BaseInstructControl baseInstructControl;


    public GetPageSourceBiz() {
    }

    @PostConstruct
    public void pageSource() {


        String headPath = "C:\\Users\\Administrator\\Desktop\\";
        String txtFileName = "kangkang.txt";
        String filePath = headPath + txtFileName;
        File file = new File(filePath);
        int lineSize = 1;
        try {
            if (file.isFile() && file.exists()) {
                InputStreamReader isr = new InputStreamReader(new FileInputStream(file), "utf-8");
                BufferedReader br = new BufferedReader(isr);
                String lineTxt = new String();
                txtReadOneLine(lineTxt, br, lineSize, txtFileName);
            } else {
                log.info("文件不存在!");
            }
        } catch (Exception e) {
            e.printStackTrace();
            log.info("文件读取错误!", e);
        }
    }

    /**
     * 主采集业务流程
     *
     * @param url
     * @param filePath
     * @param lineSize
     * @param path
     * @throws Exception
     */
    private void open(String url, String filePath, int lineSize, String path) throws Exception {

        // 创建连接
        Connection connect = Jsoup.connect(url);
        // 模拟连接头，如果不设置会被亚马逊当作爬虫拒绝访问，同时设置请求超时时间
        Document doc = connect.userAgent("Mozilla/5.0 (Windows NT 6.1; WOW64; rv:5.0) Gecko/20100101 Firefox/5.0")
                .timeout(10 * 1000).get();
        if (doc != null) {
            log.info("进入Document！！！！");
            //  根据不同类型商品反射调用不同的数据采集方法
            String commodityType = doc.select("span.a-list-item").text().replaceAll(" ", "");
            commodityType = commodityType.split("›")[0];
            if (StringUtils.isBlank(commodityType)) {
                log.info("商品类型为空 ", new ServiceException(ServiceExceptionEnum.UNSUPPORTED_ITEM_TYPE));
            }
            log.info("种类：" + commodityType);
            BaseInstruct baseInstruct = baseInstructControl.getBaseInstructControl(commodityType);
            log.info("baseInstruct = " + baseInstruct);
            baseInstruct.getPageSource(doc, filePath, lineSize, path);
        }
    }

    private void txtReadOneLine(String lineTxt, BufferedReader br, int lineSize, String txtFileName) throws Exception {
        try {
            while ((lineTxt = br.readLine()) != null) {
                log.info("当前读取到的商品url：" + lineTxt);
                String url = lineTxt.replace(" ", "");
                StringBuilder path = new StringBuilder("ys").append(DateUtil.mdFormat()).append("-").append(lineSize);
                log.info("path路径 ：" + path);
                StringBuilder pathName = new StringBuilder("C:\\Users\\Administrator\\Desktop\\").append(txtFileName.replace(".txt", "")).
                        append("\\").append(path).append("\\").append(path);
                log.info("pathName ：" + pathName);
                log.info("当前读取行数 ：" + lineSize);
                this.open(url, pathName.toString(), lineSize, path.toString());
                lineSize++;
            }
        }catch(ServiceException se){
            txtReadOneLine(lineTxt,br,lineSize,txtFileName);
        }finally {
            br.close();
        }
    }

}
