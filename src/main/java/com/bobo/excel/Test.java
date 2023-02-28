package com.bobo.excel;

import java.io.File;
/**
 * 通过excel导出  对应的数据库 的表结构
 */
public class Test {
    public static void main(String[] args) {
        DBUtil dbUtil = new DBUtil();
        dbUtil.importTableDDL(new File("F:\\360MoveData\\Users\\Administrator\\Desktop\\表结构.xlsx"));
    }
}
