package com.ys.instruct;

import org.jsoup.nodes.Document;

import java.util.List;
import java.util.Set;

/**
 * projectName : demo
 *
 * @author yangshuai
 * @version ID : V1.0 , 2018/12/1718:47
 */
public interface Attributefield {

    public Set<String> imgUrl (Document doc, Set < String > imgUrls, String txtFilePath, String path);

    public String price (Document doc);

    public List<String> colorList (Document doc, List < String > colorList);

    public List<String> sizeContainer (Document doc, List < String > sizeList);

    public List<String> productName (Document doc);

    public String salePrice (Document doc);

    public List<String> bulletPointList (Document doc, List < String > bulletPoint);

    public String itemType (Document doc, StringBuilder itemType );

}
