package com.ys.dto;

import com.ys.common.utils.DateUtil;
import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

/**
 * 服装类dto
 */

@Getter
@Setter
@ToString
public class CommodityDTO implements Serializable {

    /**
     * 反序列化
     */
    private static final long serialVersionUID = 2188991160191284648L;

    private int lineSize;

    private Set<String> imgUrls = new HashSet<>();

    private List<String> colorList = new ArrayList<>();

    private List<String> sizeList = new ArrayList<>();

    private List<String> bulletPoint = new ArrayList<>();

    private String itemName;

    private String productIDType = "UPC";

    private String brandName = "Surprise4U";

    private String fulfillmentLatency = "8";

    private String quantity = "100";

    private String salePrice;

    private String saleFromDate = new DateUtil().y_M_dFormat();

    private String saleEndDate = new DateUtil().dateYearsAlgorithm(10);

    private String shippingTemplate = "美国模板9";

    private String parentChild;

    private String parentSsku;

    private String relationshipType;

    private String variationTheme = "SizeColor";

    private String fabricType = "Cotton";

    private String importDesignation = "Imported";

    private String itemType ;

    private String standardPrice;

    private String listPrice;

    private String gender;

}
