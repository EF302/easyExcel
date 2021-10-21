package com.lwl.easyexcel.converter;


import com.alibaba.excel.converters.Converter;
import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.GlobalConfiguration;
import com.alibaba.excel.metadata.property.ExcelContentProperty;

/**
 * 性别对应数据转换
 */
public class SexConverter implements Converter<Integer> {

    public static final String MALE = "男";

    public static final String FEMALE = "女";

    /**
     * 数据库存储的数据类型
     *
     * @return
     */
    @Override
    public Class supportJavaTypeKey() {
        return Integer.class;
    }

    /**
     * excel中的数据类型
     *
     * @return
     */
    @Override
    public CellDataTypeEnum supportExcelTypeKey() {
        return CellDataTypeEnum.STRING;
    }

    /**
     * 导入：性别数据转换——男：1，女：2
     *
     * @param cellData
     * @param excelContentProperty
     * @param globalConfiguration
     * @return
     * @throws Exception
     */
    @Override
    public Integer convertToJavaData(CellData cellData, ExcelContentProperty excelContentProperty, GlobalConfiguration globalConfiguration) throws Exception {
        String stringValue = cellData.getStringValue();
        if (MALE.equals(stringValue)){
            return 1;
        }else {
            return 2;
        }
    }

    /**
     * 导出：数据转换——1：男，2：女
     *
     * @param integer
     * @param excelContentProperty
     * @param globalConfiguration
     * @return
     * @throws Exception
     */
    @Override
    public CellData convertToExcelData(Integer integer, ExcelContentProperty excelContentProperty, GlobalConfiguration globalConfiguration) throws Exception {
       if(integer.equals(1)){
           return new CellData("男");
       }else {
           return new CellData("女");
       }
    }
}
