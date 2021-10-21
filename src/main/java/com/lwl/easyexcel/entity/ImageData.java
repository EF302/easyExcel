package com.lwl.easyexcel.entity;

import java.io.File;
import java.io.InputStream;
import java.net.URL;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import com.alibaba.excel.annotation.write.style.ContentRowHeight;
import com.alibaba.excel.converters.string.StringImageConverter;
import lombok.Data;


/**
 * 图片导出的5种方式
 */
@Data
@ContentRowHeight(200)
@ColumnWidth(200 / 8)
public class ImageData {

    /**
     * 文件
     */
    @ExcelProperty("文件对象")
    private File file;

    /**
     * 输入流
     */
    @ExcelProperty("输入流对象")
    private InputStream inputStream;

    /**
     * 图片地址
     */
    @ExcelProperty(value = "string类型 ", converter = StringImageConverter.class)
    //@ExcelProperty(value = "string类型 ")
    private String string;

    /**
     * 字节数组
     */
    private byte[] byteArray;

    /**
     * 根据url导出 版本2.1.1才支持该种模式
     */
    private URL url;
}
