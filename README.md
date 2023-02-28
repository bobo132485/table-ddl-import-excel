# table-ddl-import-excel
通过excel导出mysql表结构解析

###示例
####修改 db.properties 配置文件
```java
//运行代码
public class Test {
    public static void main(String[] args) {
        DBUtil dbUtil = new DBUtil();
        dbUtil.importTableDDL(new File("F:\\360MoveData\\Users\\Administrator\\Desktop\\表结构.xlsx"));
    }
}
```
###结果
![结果](https://github.com/bobo132485/table-ddl-import-excel/blob/main/data.png)