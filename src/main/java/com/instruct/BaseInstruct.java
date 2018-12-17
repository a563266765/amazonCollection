package com.instruct;

import com.dto.CommodityDTO;
import org.jsoup.nodes.Document;

/**
 * 公共接口指令类
 * projectName : demo
 *
 * @author yangshuai
 * @version ID : V1.0 , 2018/12/1221:51
 */
public interface BaseInstruct {

    void getPageSource(Document doc, String filePath, int lineSize, String path);
}
