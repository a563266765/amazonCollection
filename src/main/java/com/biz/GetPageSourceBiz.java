package com.biz;

import com.common.utils.DateUtil;
import com.control.BaseInstructControl;
import com.instruct.BaseInstruct;
import org.jsoup.Connection;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStreamReader;

@Service
public class GetPageSourceBiz {

	@Autowired
	private CommodityInstruct excelExportServlet;

	@Autowired
	private BaseInstructControl baseInstructControl;

	public GetPageSourceBiz(){}
	public void pageSource() {

		try {
			String headPath = "C:\\Users\\Administrator\\Desktop\\";
			String txtFileName = "kangkang.txt";
			String filePath = headPath + txtFileName;
			File file = new File(filePath);
			DateUtil dateUtil = new DateUtil();
			if (file.isFile() && file.exists()) {
				InputStreamReader isr = new InputStreamReader(new FileInputStream(file), "utf-8");
				BufferedReader br = new BufferedReader(isr);
				String lineTxt;
				int lineSize = 1;
				while ((lineTxt = br.readLine()) != null) {
					String url = lineTxt.replace(" ", "");
					String path = new StringBuilder("ys").append(dateUtil.mdFormat()).append("-").append(lineSize).toString();
					String pathName = new StringBuilder("C:\\Users\\Administrator\\Desktop\\").append(txtFileName.replace(".txt", "")).
							append("\\").append(path).append("\\").append(path).toString();
					this.open(url, pathName, lineSize, path);
					lineSize++;
				}
				br.close();
			} else {
				System.out.println("文件不存在!");
			}
		} catch (Exception e) {
			System.out.println("文件读取错误!");
		}
	}

	/**
	 * 主采集业务流程
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
			System.out.println("进入Document！！！！");
			//  根据不同类型商品反射调用不同的数据采集方法
			String commodityType  = doc.select("span.nav-a-content").text().replaceAll(" ","");
			BaseInstruct baseInstruct = baseInstructControl.getBaseInstructControl(commodityType);
			baseInstruct.getPageSource(doc,filePath,lineSize,path);
		}
	}

}
