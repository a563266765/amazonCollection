package com.ys.control;

import com.ys.common.enums.CommodityTypeEnum;
import com.ys.common.exception.ServiceException;
import com.ys.common.exception.ServiceExceptionEnum;
import com.ys.instruct.BaseInstruct;
import lombok.extern.slf4j.Slf4j;
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
@Slf4j
public class BaseInstructControl {

    @Autowired
    private BaseInstruct commodityInstruct;

    @Autowired
    private BaseInstruct petSuppliesInstruct;

    private Map<String,BaseInstruct> baseInstructMap = new HashMap<>();

    /**
     * 初始化方法
     */
    @PostConstruct
    public void init(){
        log.info("指令类初始化！！！");
        baseInstructMap.put(CommodityTypeEnum.CLOTHING.getCode(),commodityInstruct);
        baseInstructMap.put(CommodityTypeEnum.PET_SUPPLIES.getCode(),petSuppliesInstruct);
    }

    /**
     * 获取指令
     * @param commodityType
     * @return
     */
    public BaseInstruct getBaseInstructControl(String commodityType){
        log.info("进入获取指令类！！！");
        BaseInstruct baseInstruct = this.baseInstructMap.get(commodityType);
        if (baseInstruct == null){
            log.info("传入的商品类型有误 ：",new ServiceException(ServiceExceptionEnum.COMMODITY_TYPE_ERROR));
            throw new ServiceException(ServiceExceptionEnum.COMMODITY_TYPE_ERROR);
        }
        log.info("baseInstruct = " + baseInstruct.getClass());
        return baseInstruct;
    }
}
