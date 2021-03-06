package com.ys.common.enums;

import com.ys.common.exception.ServiceException;
import com.ys.common.exception.ServiceExceptionEnum;
import lombok.Getter;

/**
 * projectName : demo
 *
 * @author yangshuai
 * @version ID : V1.0 , 2018/12/1221:58
 */
@Getter
public enum CommodityTypeEnum {

    /** 服装类 */
    CLOTHING("Clothing,Shoes&Jewelry","服装类"),

    /** 宠物用品 */
    PET_SUPPLIES("PETSUPPLIES","宠物用品");

    private String code;

    private String values;

    CommodityTypeEnum(String code, String values){
        this.code   = code;
        this.values = values;
    }

    public CommodityTypeEnum getCommodityTypeEnum(String code){
        CommodityTypeEnum commodityTypeEnum = null;
        for (CommodityTypeEnum commodityType: CommodityTypeEnum.values()) {
            if (commodityType.equals(code)){
                commodityTypeEnum = commodityType;
                break;
            }
        }
        if (commodityTypeEnum == null){
            throw new ServiceException(ServiceExceptionEnum.COMMODITY_TYPE_ERROR);
        }
        return commodityTypeEnum;
    }
}
