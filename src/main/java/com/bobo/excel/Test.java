package com.bobo.excel;

import cn.hutool.core.collection.ListUtil;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;
import cn.hutool.poi.excel.cell.CellSetter;
import org.apache.poi.ss.usermodel.*;

import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;

/**
 * 通过excel导出  对应的数据库 的表结构
 */
public class Test {
    public static void main(String[] args) {
        DBUtil dbUtil = new DBUtil();
        //对应的数据库
        String table_schema = "xiuhu_dev";
        //获取所有的表名
        ResultSet resultSet = dbUtil.select("SELECT\n" +
                        " DISTINCT\n" +
                        "\ta.table_name ,\n" +
                        "\tb.table_COMMENT 表备注\n" +
                        "FROM\n" +
                        "\tINFORMATION_SCHEMA.COLUMNS  a\n" +
                        "\tLEFT JOIN information_schema.TABLES b ON a.TABLE_NAME = b.TABLE_NAME \n" +
                        "\tAND a.table_schema = b.table_schema\n" +
                        "WHERE\n" +
                        "a.table_schema = '"+ table_schema +"'",
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
        ExcelWriter writer = ExcelUtil.getWriter("C:\\Users\\Administrator\\Desktop\\表结构.xlsx");
        List<Object> all = new ArrayList<>();
        int i = 1;
        for (List<String> item : tables) {
            item.set(0, i + "," + item.get(0));
            all.add(item);
            all.add(buildTitle(writer, columnName));
            all.addAll(rows(dbUtil, columnName, table_schema, item.get(1)));
            all.add(ListUtil.of());
            all.add(ListUtil.of());
            i++;
        }


        writer.write(all);
        writer.flush();
        writer.close();
        dbUtil.closeConnection();
    }
    /**
     * 获取列信息
     * @param dbUtil
     * @param table_schema
     * @param columnName
     * @param tableName
     * @return
     */
    static List<List<String>> rows(DBUtil dbUtil, List<String> columnName,String table_schema,String tableName){
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
    static class Title implements CellSetter  {
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
