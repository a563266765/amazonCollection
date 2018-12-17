package com.control;

import com.common.enums.CommodityTypeEnum;
import com.common.exception.ServiceException;
import com.common.exception.ServiceExceptionEnum;
import com.instruct.BaseInstruct;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import javax.annotation.PostConstruct;
import java.util.HashMap;
import java.util.Map;

/**
 * projectName : demo
 *
 * @author yangshuai
 * @version ID : V1.0 , 2018/12/1221:53
 */
@Component
public class BaseInstructControl {

    @Autowired
    private BaseInstruct CommodityInstruct;

    @Autowired
    private BaseInstruct PetSuppliesInstruct;

    private Map<String,BaseInstruct> baseInstructMap = new HashMap<>();

    /**
     * 初始化方法
     */
    @PostConstruct
    public void init(){
        baseInstructMap.put(CommodityTypeEnum.CLOTHING.getCode(),CommodityInstruct);
        baseInstructMap.put(CommodityTypeEnum.PET_SUPPLIES.getCode(),PetSuppliesInstruct);
    }

    /**
     * 获取指令
     * @param commodityType
     * @return
     */
    public BaseInstruct getBaseInstructControl(String commodityType){

        BaseInstruct baseInstruct = this.baseInstructMap.get(commodityType);
        if (baseInstruct == null){
            throw new ServiceException(ServiceExceptionEnum.COMMODITY_TYPE_ERROR);
        }
        return baseInstruct;
    }
}
