package com.bobo.excel;

import cn.hutool.core.collection.ListUtil;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;
import cn.hutool.poi.excel.cell.CellSetter;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.sql.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;

/**
 * @author bobo
 * @version 1.0
 * @description: TODO
 * @date 2022/8/25/025 16:34
 */
public class DBUtil {

    //连接信息
    private static String driverName;
    private static String url;
    private static String username;
    private static String password;
    /** 数据库名称*/
    private static String tableSchema;

    //注册驱动，使用静态块，只需注册一次
    static {
        //初始化连接信息
        Properties properties = new Properties();
        try {
            properties.load(new FileReader("src/db.properties"));
            driverName = properties.getProperty("driverName");
            url = properties.getProperty("url");
            username = properties.getProperty("username");
            password = properties.getProperty("password");
            tableSchema = properties.getProperty("tableSchema");
        } catch (IOException e) {
            e.printStackTrace();
        }
        //1、注册驱动
        try {
            //通过反射，注册驱动
            Class.forName(driverName);
        } catch (ClassNotFoundException e) {
            e.printStackTrace();
        }
    }

    //jdbc对象
    private Connection connection = null;
    private PreparedStatement preparedStatement = null;
    private ResultSet resultSet = null;

    //获取连接
    public void getConnection() {
        try {
            //2、建立连接
            connection = DriverManager.getConnection(url, username, password);
        } catch (SQLException e) {
            e.printStackTrace();
        }
    }

    //更新操作：增删改
    public int update(String sql, Object[] objs) {
        int i = 0;
        try {
            getConnection();
            //3、创建sql对象
            preparedStatement = connection.prepareStatement(sql);
            for (int j = 0; j < objs.length; j++) {
                preparedStatement.setObject(j + 1, objs[j]);
            }
            //4、执行sql，返回改变的行数
            i = preparedStatement.executeUpdate();
        } catch (SQLException e) {
            e.printStackTrace();
        }
        return i;
    }

    //查询操作
    public ResultSet select(String sql, Object[] objs) {
        try {
            getConnection();
            //3、创建sql对象
            preparedStatement = connection.prepareStatement(sql);
            for (int j = 0; j < objs.length; j++) {
                preparedStatement.setObject(j + 1, objs[j]);
            }
            //4、执行sql，返回查询到的set集合
            resultSet = preparedStatement.executeQuery();
        } catch (SQLException e) {
            e.printStackTrace();
        }
        return resultSet;
    }

    //断开连接
    public void closeConnection() {
        //5、断开连接
        if (resultSet != null) {
            try {
                resultSet.close();
            } catch (SQLException e) {
                e.printStackTrace();
            }
        }
        if (preparedStatement != null) {
            try {
                preparedStatement.close();
            } catch (SQLException e) {
                e.printStackTrace();
            }
        }
        if (connection != null) {
            try {
                connection.close();
            } catch (SQLException e) {
                e.printStackTrace();
            }
        }
    }

    public void importTableDDL(File file){
        //获取所有的表名
        ResultSet resultSet = select("SELECT\n" +
                        " DISTINCT\n" +
                        "\ta.table_name ,\n" +
                        "\tb.table_COMMENT 表备注\n" +
                        "FROM\n" +
                        "\tINFORMATION_SCHEMA.COLUMNS  a\n" +
                        "\tLEFT JOIN information_schema.TABLES b ON a.TABLE_NAME = b.TABLE_NAME \n" +
                        "\tAND a.table_schema = b.table_schema\n" +
                        "WHERE\n" +
                        "a.table_schema = '"+ tableSchema +"'",
                new Object[]{});
        List<List<String>> tables = new ArrayList<>();
        try {
            while (resultSet.next()) {
                List<String> list = new ArrayList<>();
                list.add(resultSet.getString("表备注"));
                list.add(resultSet.getString("table_name"));
                tables.add(list);
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }
        List<String> columnName = ListUtil.of("序号", "列名", "数据类型", "长度", "主键", "允许空", "备注");
        ExcelWriter writer = ExcelUtil.getWriter(file);
        List<Object> all = new ArrayList<>();
        int i = 1;
        for (List<String> item : tables) {
            item.set(0, i + "," + item.get(0));
            all.add(item);
            all.add(buildTitle(writer, columnName));
            all.addAll(rows(this, columnName, tableSchema, item.get(1)));
            all.add(ListUtil.of());
            all.add(ListUtil.of());
            i++;
        }


        writer.write(all);
        writer.flush();
        writer.close();
        closeConnection();
    }


    /**
     * 获取列信息
     * @param dbUtil
     * @param table_schema
     * @param columnName
     * @param tableName
     * @return
     */
    static List<List<String>> rows(DBUtil dbUtil, List<String> columnName, String table_schema, String tableName){
        List<List<String>> list = new ArrayList<>();
        Object[] objs = {tableName};
        ResultSet resultSet = dbUtil.select("SELECT\n" +
                "@i:=@i+1 AS '序号',\n" +
                "COLUMN_NAME 列名,\n" +
                "DATA_TYPE 数据类型,\n" +
                "CHARACTER_MAXIMUM_LENGTH 长度,\n" +
                "if(Column_key = 'PRI', 'Y', '') 主键,\n" +
                "IS_NULLABLE 允许空,\n" +
                "COLUMN_COMMENT 备注 \n" +
                "FROM\n" +
                "INFORMATION_SCHEMA.COLUMNS, (SELECT @i:=0) AS itable\n" +
                "WHERE\n" +
                "table_schema = '" + table_schema + "'\n" +
                "and table_name = ?", objs);
        try {
            while (resultSet.next()) {
                List<String> row = new ArrayList<>();
                for (String s : columnName) {
                    row.add(resultSet.getString(s));
                }
                list.add(row);
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }
        return list;
    }

    static List<Title> buildTitle(ExcelWriter excelWriter, List<String> columnName){
        List<Title> list = new ArrayList<>();
        for (String s : columnName) {
            list.add(new Title(s, excelWriter));
        }
        return list;
    }

    /**
     * 美化标题
     */
    static class Title implements CellSetter {
        private String value;
        private ExcelWriter excelWriter;
        Title(String value, ExcelWriter excelWriter){
            this.value = value;
            this.excelWriter = excelWriter;
        }
        @Override
        public void setValue(Cell cell) {
            cell.setCellValue(value);
            //设置单元格的颜色字体
            CellStyle style = excelWriter.createCellStyle();
            style.setFillForegroundColor(IndexedColors.AQUA.getIndex());
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            style.setVerticalAlignment(VerticalAlignment.CENTER); //设置垂直居中
            style.setBorderBottom(BorderStyle.THIN);
            style.setBorderLeft(BorderStyle.THIN); // 设置边界的类型单元格的左边框
            style.setBorderRight(BorderStyle.THIN);
            style.setBorderTop(BorderStyle.THIN);
            Font font = excelWriter.createFont();
            //默认字体为宋体
            font.setFontName("宋体");
            //设置字体大小
            font.setFontHeightInPoints((short) 14); //设置字号
            //设置字体颜色
            font.setColor(IndexedColors.WHITE.getIndex());
            //设置字体加粗
            font.setBold(true);
            style.setFont(font);
            cell.setCellStyle(style);
        }
    }
}
