package com.lwl.easyexcel.test;

import java.io.File;
import java.net.URL;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.alibaba.excel.util.FileUtils;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.fill.FillConfig;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import com.alibaba.excel.write.style.column.LongestMatchColumnWidthStyleStrategy;
import com.alibaba.fastjson.JSON;
import com.lwl.easyexcel.entity.FillData;
import com.lwl.easyexcel.entity.ImageData;
import com.lwl.easyexcel.entity.User;
import com.lwl.easyexcel.listener.ExcelReadListener;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.junit.Test;

/**
 * todo
 *
 * @author longwanli
 * @date 2021/7/8 11:13
 */
@Slf4j
public class TestExcel {

    //===============================================导入==============================

    /**
     * 导入
     */
    @Test
    public void read() {
        String filePath = "F:\\IDEAWorkSpace\\study2021\\easyExcel\\src\\main\\java\\com\\lwl\\easyexcel\\excel\\user.xlsx";
        //read()有多种构造函数，也可接收文件对象、文件流。
        //sheet()用于指定第几个表单，默认是第一个,有多种构造。
        //headRowNumber()指定excel中表头占多少行
        EasyExcel.read(filePath, User.class, new ExcelReadListener()).sheet().headRowNumber(1).doRead();

        ////写法2 注 需要手动关闭流
        ////ExcelReader excelReader = EasyExcel.read(filePath, User.class, new ExcelReadListener()).build();
        ////ReadSheet readSheet = EasyExcel.readSheet(0).build();
        ////或者
        //ExcelReader excelReader = EasyExcel.read(filePath).build();
        //ReadSheet readSheet = EasyExcel.readSheet(0).head(User.class).registerReadListener(new ExcelReadListener()).build();
        //excelReader.read(readSheet);
        //excelReader.finish();
    }

    /**
     * 导入：读取多个、部分sheet
     */
    @Test
    public void read2() {
        String filePath = "F:\\IDEAWorkSpace\\study2021\\easyExcel\\src\\main\\java\\com\\lwl\\easyexcel\\excel\\user.xlsx";
        // 方法1：doReadAll()
        //EasyExcel.read(filePath, User.class, new ExcelReadListener()).doReadAll();

        // 方法2： 注 需要手动关闭流
        ExcelReader excelReader = EasyExcel.read(filePath).build();
        // readSheet参数设置读取sheet的序号  这里为了简单，所以注册了同样的head和Listener
        ReadSheet readSheet1 = EasyExcel.readSheet(0).head(User.class).registerReadListener(new ExcelReadListener()).build();
        ReadSheet readSheet2 = EasyExcel.readSheet(1).head(User.class).registerReadListener(new ExcelReadListener()).build();
        // 注意  一定要把sheet1、sheet2一起传进去，不然有个问题，03版的excel会读取多次，浪费性能
        excelReader.read(readSheet1, readSheet2);
        excelReader.finish();
    }

    //=========================================导出============================

    /**
     * 导出
     */
    @Test
    public void write() {
        String filePath = "F:\\IDEAWorkSpace\\study2021\\easyExcel\\src\\main\\java\\com\\lwl\\easyexcel\\excel\\user2.xlsx";
        User user1 = User.builder().name("zs").age(19).sex(1).birthday("2021-07-08 00:00:00").height("172.25").remark("我是一个人类").build();
        User user2 = User.builder().name("ls").age(18).sex(1).birthday("2021-07-08 00:00:00").height("172.56").remark("我是一个人类").build();
        User user3 = User.builder().name("ww").age(19).sex(1).birthday("2021-07-08 00:00:00").height("172.0").remark("我是一个人类").build();
        User user4 = User.builder().name("ch").age(20).sex(2).birthday("2021-07-08 00:00:00").height("172").remark("我是一个人类").build();
        List<User> userList = new ArrayList<>();
        userList.add(user1);
        userList.add(user2);
        userList.add(user3);
        userList.add(user4);

        //自定义动态表头
        List<List<String>> tableHead = new ArrayList<>();
        tableHead.add(Arrays.asList("姓名"));
        tableHead.add(Arrays.asList("年龄"));
        tableHead.add(Arrays.asList("性别"));
        tableHead.add(Arrays.asList("生日"));
        tableHead.add(Arrays.asList("身高"));
        //tableHead.add(Arrays.asList("备注"));

        //方法1：
        EasyExcel.write(filePath, User.class).sheet("用户").doWrite(userList);
        //==============测试自定义动态表头
        //EasyExcel.write(filePath, User.class).head(tableHead).sheet("用户").doWrite(userList);
        //等效
        EasyExcel.write(filePath).head(tableHead).sheet("用户").doWrite(userList);
        //===========自动列宽测试：结果不理想
        // 注 POI对中文的自动列宽适配不友好，easyexcel对数字也不能准确适配列宽，虽然提供的适配策略，但也不能精确适配，可以自己重写
        //LongestMatchColumnWidthStyleStrategy longestMatchColumnWidthStyleStrategy = new LongestMatchColumnWidthStyleStrategy();
        //EasyExcel.write(filePath).registerWriteHandler(longestMatchColumnWidthStyleStrategy).sheet("用户").doWrite(userList);
        //==============合并单元格测试
        //第1列，每两行合并为一个单元格 注 如果两个单元格都有数据，合并后，数据为第一个单元格的数据
        //LoopMergeStrategy loopMergeStrategy = new LoopMergeStrategy(2, 0);
        //EasyExcel.write(filePath, User.class).registerWriteHandler(loopMergeStrategy).sheet("用户").doWrite(userList);


        //方法2：注 需要手动关闭流
        //// 这里 需要指定写用哪个class去写
        //ExcelWriter excelWriter = EasyExcel.write(filePath, User.class).head(tableHead).build();
        //WriteSheet writeSheet = EasyExcel.writerSheet("用户").build();
        //excelWriter.write(userList, writeSheet);
        ////千万别忘记finish关闭流
        //excelWriter.finish();

        //注 超大量数据导出  主要思想：分批获取数据导出，如：分页查询
        //ExcelWriter excelWriter = EasyExcel.write(filePath, User.class).head(tableHead).build();
        //WriteSheet writeSheet = EasyExcel.writerSheet("用户").build();
        //Integer pageSize = 10000;
        ////注 总页数可通过一次分页查询得的
        //Integer pageTotal = null;
        //List<User> userList2 = null;
        //for (Integer i = 0; i < pageTotal; i+=pageSize) {
        //    //userList2 = userService.list(new Page(i, pageSize))
        //    excelWriter.write(userList2, writeSheet);
        //    userList2.clear();
        //}
        ////千万别忘记finish关闭流
        //excelWriter.finish();
    }

    /**
     * 自定义样式：设置头策略、内容策略
     */
    @Test
    public void styleWrite() {
        String filePath = "F:\\IDEAWorkSpace\\study2021\\easyExcel\\src\\main\\java\\com\\lwl\\easyexcel\\excel\\user2.xlsx";
        User user1 = User.builder().name("zs").age(19).sex(1).birthday("2021-07-08 00:00:00").height("172.25").remark("我是一个人类").build();
        User user2 = User.builder().name("ls").age(18).sex(1).birthday("2021-07-08 00:00:00").height("172.56").remark("我是一个人类").build();
        User user3 = User.builder().name("ww").age(19).sex(1).birthday("2021-07-08 00:00:00").height("172.0").remark("我是一个人类").build();
        User user4 = User.builder().name("ch").age(20).sex(2).birthday("2021-07-08 00:00:00").height("172").remark("我是一个人类").build();
        List<User> userList = Arrays.asList(user1, user2, user3, user4);

        //自定义动态表头
        List<List<String>> head = new ArrayList<>();
        head.add(Arrays.asList("姓名"));
        head.add(Arrays.asList("年龄"));
        head.add(Arrays.asList("性别"));
        head.add(Arrays.asList("生日"));
        head.add(Arrays.asList("身高"));
        head.add(Arrays.asList("备注"));

        // 头的策略
        WriteCellStyle headWriteCellStyle = new WriteCellStyle();
        //设置表头居中对齐
        headWriteCellStyle.setHorizontalAlignment(HorizontalAlignment.JUSTIFY);
        // 背景设置为红色
        headWriteCellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
        // 设置头字体粗细
        WriteFont headWriteFont = new WriteFont();
        headWriteFont.setFontHeightInPoints((short) 10);
        headWriteCellStyle.setWriteFont(headWriteFont);

        // 内容的策略
        WriteCellStyle contentWriteCellStyle = new WriteCellStyle();
        //设置内容靠左对齐
        contentWriteCellStyle.setHorizontalAlignment(HorizontalAlignment.LEFT);
        // 注 需要指定FillPatternType为FillPatternType.SOLID_FOREGROUND，不然无法显示背景颜色。
        contentWriteCellStyle.setFillPatternType(FillPatternType.SOLID_FOREGROUND);
        // 背景色
        contentWriteCellStyle.setFillForegroundColor(IndexedColors.DARK_YELLOW.getIndex());
        // 字体及大小
        WriteFont contentWriteFont = new WriteFont();
        contentWriteFont.setFontName("宋体");
        contentWriteFont.setFontHeightInPoints((short) 8);
        contentWriteCellStyle.setWriteFont(contentWriteFont);

        // 注 头的样式、内容的样式设置是分离的，故需要分别设置并注入
        HorizontalCellStyleStrategy horizontalCellStyleStrategy =
                new HorizontalCellStyleStrategy(headWriteCellStyle, contentWriteCellStyle);
        //自动列宽
        LongestMatchColumnWidthStyleStrategy longestMatchColumnWidthStyleStrategy = new LongestMatchColumnWidthStyleStrategy();
        //导出
        EasyExcel.write(filePath, User.class).registerWriteHandler(horizontalCellStyleStrategy).sheet("用户").doWrite(userList);
        //注 可以注册多个registerWriteHandler
        // 注意 不知道为什么，如果用自定义的动态表头，内容设置的背景颜色没有效果  待弄
        //EasyExcel.write(filePath).head(head).registerWriteHandler(horizontalCellStyleStrategy).registerWriteHandler(longestMatchColumnWidthStyleStrategy).sheet("用户").doWrite(userList);

    }

    /**
     * 不导出、只导出指定的列
     */
    @Test
    public void write2() {
        String filePath = "F:\\IDEAWorkSpace\\study2021\\easyExcel\\src\\main\\java\\com\\lwl\\easyexcel\\excel\\user2.xlsx";
        User user1 = User.builder().name("zs").age(19).sex(1).birthday("2021-07-08 00:00:00").height("172.25").remark("我是一个人类").build();
        User user2 = User.builder().name("ls").age(18).sex(1).birthday("2021-07-08 00:00:00").height("172.56").remark("我是一个人类").build();
        User user3 = User.builder().name("ww").age(19).sex(1).birthday("2021-07-08 00:00:00").height("172.0").remark("我是一个人类").build();
        User user4 = User.builder().name("ch").age(20).sex(2).birthday("2021-07-08 00:00:00").height("172").remark("我是一个人类").build();
        List<User> userList = new ArrayList<>();
        userList.add(user1);
        userList.add(user2);
        userList.add(user3);
        userList.add(user4);

        //二级头设置
        List<List<String>> head = new ArrayList<>();
        head.add(Arrays.asList("主标题", "姓名"));
        head.add(Arrays.asList("主标题", "年龄"));
        head.add(Arrays.asList("主标题", "性别"));
        head.add(Arrays.asList("主标题", "生日"));
        head.add(Arrays.asList("主标题", "身高"));

        // 不导出name
        //Set<String> excludeColumnFiledNames = new HashSet<String>();
        //excludeColumnFiledNames.add("name");
        //EasyExcel.write(filePath, User.class)
        //        .excludeColumnFiledNames(excludeColumnFiledNames)
        //        .head(head)
        //        .sheet("用户")
        //        .doWrite(userList);

        // 只导出name
        Set<String> includeColumnFiledNames = new HashSet<String>();
        includeColumnFiledNames.add("name");
        EasyExcel.write(filePath, User.class)
                .includeColumnFiledNames(includeColumnFiledNames)
                .head(head)
                .sheet("用户")
                .doWrite(userList);
    }

    /**
     * 图片导出
     *
     * @throws Exception
     */
    @Test
    public void imageWrite() throws Exception {
        String fileName = "F:\\IDEAWorkSpace\\study2021\\easyExcel\\src\\main\\java\\com\\lwl\\easyexcel\\excel\\image.xlsx";
        ImageData imageData = null;
        try {
            List<ImageData> imageDataList = new ArrayList<>();
            imageData = new ImageData();
            String imagePath = "F:\\IDEAWorkSpace\\study2021\\easyExcel\\src\\main\\java\\com\\lwl\\easyexcel\\excel\\宝箱钥匙.jpg";
            // 放入五种类型的图片 根据实际使用只要选一种即可
            //注 以下四种方式只适用于导出本地服务器的图片
            imageData.setByteArray(FileUtils.readFileToByteArray(new File(imagePath)));
            imageData.setFile(new File(imagePath));
            imageData.setString(imagePath);
            imageData.setInputStream(FileUtils.openInputStream(new File(imagePath)));
            //注 URL方式可实现通过地址访问远端服务器的图片，并导出到excel中
            imageData.setUrl(new URL("http://hongmofang-test.obs.cn-north-4.myhuaweicloud.com/v5mgr-Backend/1621953354251.jpeg"));
            imageDataList.add(imageData);

            EasyExcel.write(fileName, ImageData.class).sheet().doWrite(imageDataList);
        } finally {
            if (imageData != null && imageData.getInputStream() != null) {
                imageData.getInputStream().close();
            }
            log.info("导出完成！");
        }
    }


    /**
     * 根据模板导出
     */
    @Test
    public void simpleFill() {
        // 模板 注意 用{}来表示要用的变量 如果本来就有特殊字符{ 、}，则用"\{"、"\}"代替
        String templateFileName = "F:\\IDEAWorkSpace\\study2021\\easyExcel\\src\\main\\java\\com\\lwl\\easyexcel\\excel\\template.xlsx";
        // 实例
        String fileName = "F:\\IDEAWorkSpace\\study2021\\easyExcel\\src\\main\\java\\com\\lwl\\easyexcel\\excel\\templateExample.xlsx";

        // 方案1：根据对象填充，填充到第一个sheet， 然后文件流会自动关闭
        //FillData fillData = new FillData("丁春秋", 25);
        //EasyExcel.write(fileName).withTemplate(templateFileName).sheet().doFill(fillData);

        // 方案2：根据Map填充
        //Map<String, Object> map = new HashMap<String, Object>();
        //map.put("name", "李春秋");
        //map.put("age", 21);
        //EasyExcel.write(fileName).withTemplate(templateFileName).sheet().doFill(map);

        //===========多行数据填充 注 上文方案1、2只能填充一行
        // 模板 注意 此时模板中变量引用为：｛.name｝，点表示该参数是集合
        String templateFileName2 = "F:\\IDEAWorkSpace\\test\\src\\main\\java\\com\\lwl\\test\\测试2_7月\\easyExcel\\excel\\template2.xlsx";
        ExcelWriter excelWriter = EasyExcel.write(fileName).withTemplate(templateFileName2).build();
        WriteSheet writeSheet = EasyExcel.writerSheet().build();
        // 注 forceNewRow（）代表在写入list的时，不管list下面有没有空行都会创建一行，然后下面的数据往后移动。默认是false，会直接使用下一行，如果没有，则创建。如果设置为true,有个缺点：会把所有的数据都放到内存，慎用。简单的说就是，如果模板有list，且list不是最后一行，即还有数据需要填充，就必须设置 forceNewRow=true ，但是这个就会把所有数据放到内存 会很耗内存
        FillConfig fillConfig = FillConfig.builder().forceNewRow(Boolean.TRUE).build();

        List<FillData> fillDataList = Arrays.asList(
                new FillData("丁春秋", 20),
                new FillData("李夏冬", 18));
        excelWriter.fill(fillDataList, fillConfig, writeSheet);

        List<FillData> fillDataList2 = Arrays.asList(
                new FillData("张三", 19),
                new FillData("李四", 28));
        excelWriter.fill(fillDataList2, fillConfig, writeSheet);

        // 非集合注入数据
        FillData fillData = new FillData("冉菜菜", 25);
        excelWriter.fill(fillData, fillConfig, writeSheet);
        excelWriter.finish();
    }

    /**
     * 同步的返回 注 不推荐使用，如果数据量大，会把数据放到内存里面
     */
    @Test
    public void synchronousRead() {
        String fileName = "F:\\IDEAWorkSpace\\study2021\\easyExcel\\src\\main\\java\\com\\lwl\\easyexcel\\excel\\user.xlsx";

        //方法1：指定读用哪个.class去读
        List<Object> list = EasyExcel.read(fileName).head(User.class).sheet().doReadSync();
        User user;
        for (Object obj : list) {
            //返回每条数据的键值对为：映射类属性名和数据值
            user = (User) obj;
            System.out.println("读取到数据: " + JSON.toJSONString(user));
        }

        //方法2：不指定class
        // 这里也可以不指定class，返回一个list，然后读取第一个sheet，同步读取会自动finish
        List<Object> list2 = EasyExcel.read(fileName).sheet().doReadSync();
        Map<Integer, Object> map;
        for (Object obj : list2) {
            // 返回每条数据的键值对为：数据所在的列和数据值
            map = (Map<Integer, Object>) obj;
            System.out.println("读取到数据: " + JSON.toJSONString(map));
        }
    }


}
