package com.lwl.easyexcel.entity;

import java.io.Serializable;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.format.DateTimeFormat;
import com.alibaba.excel.annotation.format.NumberFormat;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import com.alibaba.excel.annotation.write.style.ContentRowHeight;
import com.alibaba.excel.annotation.write.style.HeadRowHeight;
import com.lwl.easyexcel.converter.SexConverter;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
//设置表头高度
@HeadRowHeight(20)
//设置数据单元格高度
@ContentRowHeight(10)
//设置单元格列宽
@ColumnWidth(10)
public class User implements Serializable {

    /**
     * 姓名
     */
    @ExcelProperty(index = 0)// 0代表第一列，以此类推.
    private String name;

    /**
     * 年龄
     */
    @ExcelProperty(index = 1)
    private Integer age;

    /**
     * 性别: 1对应男，2对应女
     */
    @ExcelProperty(value = "性别", converter = SexConverter.class)
    private Integer sex;

    /**
     * 生日
     */
    @ExcelProperty(value = "生日", index = 3)
    //格式化日期
    @DateTimeFormat("yyyy-MM-dd HH-mm-ss")
    @ColumnWidth(30)
    private String birthday;

    /**
     * 身高
     */
    @ExcelProperty(value = "身高", index = 4)
    //格式化数字
    @NumberFormat("###.#")
    private String height;

    /**
     * 备注
     */
    @ExcelProperty("备注")
    private String remark;

    /**
     * 头像
     */
    @ExcelProperty(value = "头像")
    private String url;
}

