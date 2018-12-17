package com.common.exception;

import lombok.Getter;
import lombok.Setter;

/**
 * projectName : demo
 *
 * @author yangshuai
 * @version ID : V1.0 , 2018/12/1311:14
 */
@Getter
@Setter
public class ServiceException extends RuntimeException {

    private final ServiceExceptionEnum errorCode;

    public ServiceException(ServiceExceptionEnum errorCode){
        super(errorCode.getMessage());
        this.errorCode = errorCode;
    }

    public ServiceException(ServiceExceptionEnum errorCode, String errorMessage){
        super(errorMessage);
        this.errorCode = errorCode;
    }
}
