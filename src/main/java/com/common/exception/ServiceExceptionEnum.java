package com.common.exception;

import lombok.Getter;
import org.apache.commons.lang.StringUtils;

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
    COMMODITY_TYPE_ERROR("COMMODITY_TYPE_ERROR","商品类型参数有误");

    private String code;
    private String message;

    ServiceExceptionEnum(String code, String message){
        this.code   = code;
        this.message = message;
    }
}
