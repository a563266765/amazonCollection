package com.ys.common.exception;

import lombok.Getter;

/**
 * projectName : demo
 *
 * @author yangshuai
 * @version ID : V1.0 , 2018/12/1310:59
 */
@Getter
public enum ServiceExceptionEnum {

    /**
     * 商品类型参数有误
     */
    COMMODITY_TYPE_ERROR("COMMODITY_TYPE_ERROR","商品类型参数有误"),

    /**
     * 商品描述取值有误
     */
    BULLET_POINT_ERROR("BULLET_POINT_ERROR","商品描述取值有误"),

    /**
     * 不支持采集的商品类型
     */
    UNSUPPORTED_ITEM_TYPE("UNSUPPORTED_ITEM_TYPE","不支持采集的商品类型");

    private String code;
    private String message;

    ServiceExceptionEnum(String code, String message){
        this.code   = code;
        this.message = message;
    }
}
