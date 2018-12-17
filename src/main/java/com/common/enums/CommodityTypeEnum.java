package com.common.enums;

import com.common.exception.ServiceException;
import com.common.exception.ServiceExceptionEnum;
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
    CLOTHING("CLOTHING","CommodityBiz"),

    /** 宠物用品 */
    PET_SUPPLIES("PETSUPPLIES","PetSuppliesBiz");

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
