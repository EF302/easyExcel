package com.lwl.easyexcel.entity;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * todo
 *
 * @author longwanli
 * @date 2021/7/9 13:24
 */
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class FillData {

    private String name;

    private double age;
}

