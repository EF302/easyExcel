package com.lwl.easyexcel.controller;


import java.io.IOException;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.List;

import javax.servlet.http.HttpServletResponse;

import com.alibaba.excel.EasyExcel;
import com.lwl.easyexcel.entity.User;
import com.lwl.easyexcel.listener.ExcelReadListener;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

/**
 * todo
 *
 * @author longwanli
 * @date 2021/7/12 9:04
 */
@RestController
@RequestMapping("/excel/")
public class ExcelController {

    /**
     * 导入
     * 注意 包名包含中文会报错，测试时改一下包名：create asm serializer error, verson 1.2.76, class class com.lwl.test.测试2_7月.easyExcel.entity.User
     *
     * @param file
     * @return
     * @throws IOException
     */
    @PostMapping("upload")
    public String upload(MultipartFile file) throws IOException {
        EasyExcel.read(file.getInputStream(), User.class, new ExcelReadListener()).sheet().doRead();
        return "SUCCESS";
    }

    /**
     * 导出
     *
     * @param response
     * @throws IOException
     */
    @GetMapping("download")
    public void download(HttpServletResponse response) throws IOException {
        //注 构造数据
        User user1 = User.builder().name("zs").age(19).sex(1).birthday("2021-07-08 00:00:00").height("172.25").remark("我是一个人类").build();
        User user2 = User.builder().name("ls").age(18).sex(1).birthday("2021-07-08 00:00:00").height("172.56").remark("我是一个人类").build();
        User user3 = User.builder().name("ww").age(19).sex(1).birthday("2021-07-08 00:00:00").height("172.0").remark("我是一个人类").build();
        User user4 = User.builder().name("ch").age(20).sex(2).birthday("2021-07-08 00:00:00").height("172").remark("我是一个人类").build();
        List<User> userList = new ArrayList<>();
        userList.add(user1);
        userList.add(user2);
        userList.add(user3);
        userList.add(user4);

        response.setContentType("application/vnd.ms-excel");
        response.setCharacterEncoding("utf-8");
        // 这里URLEncoder.encode可以防止中文乱码 当然和easyexcel没有关系
        String fileName = URLEncoder.encode("数据写出", "UTF-8");
        response.setHeader("Content-disposition", "attachment;filename=" + fileName + ".xlsx");
        EasyExcel.write(response.getOutputStream(), User.class).sheet("模板").doWrite(userList);
    }


}
