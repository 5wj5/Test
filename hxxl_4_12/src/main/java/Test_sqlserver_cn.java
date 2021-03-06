import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class Test_sqlserver_cn {

    public static int count = 0;
    public static int nullCount = 0;
    public static int nameCount = 0;
    public static String tableName = null;

    public static void main(String[] args) throws IOException {
        InputStream input = Test_sqlserver_cn.class.getClassLoader().getResourceAsStream("123.xlsx");
        Workbook wb = new XSSFWorkbook(input);
        Sheet sheet = wb.getSheet("Sheet8");
        List<Row> rows = new ArrayList<Row>();
        int physicalNumberOfRows = sheet.getPhysicalNumberOfRows();

        int rowNumberBefore = 0;
        int rowNumberAfter = 0;

        String sql = "";
//        空标志
        String nullSql = "";
//        类型标志
        String classSql = "";
//        注释标志
        String comment = "";

        String nullFlag = "";

        String nullOracleFlag = "";

        boolean flag = true;
//        遍历每一行，比较第二列的值，找到后返回
        for (int i = 0; i < physicalNumberOfRows; ++i) {
            Row row = sheet.getRow(i);

            if (row.getCell(0) != null) {
                row.getCell(0).setCellType(Cell.CELL_TYPE_STRING);

            }

            if ("1".equals(row.getCell(0).getStringCellValue().replace(" ", ""))) {
                flag = false;
                continue;
            }

            if ("0".equals(row.getCell(0).getStringCellValue().replace(" ", ""))) {

//                在这做一个剩余处理
                System.out.println("打印需要删除的的sql：");
                for (Row row2 : rows) {
//                    字段名
                    String value = row2.getCell(0).getStringCellValue().replace(" ", "");
//                    去掉[]
                    if (value.contains("[")) {
                        value = value.replace("[", "").replace("]", "");
                    }
                    sql = "alert table " + tableName + " drop column " + value;
                    System.out.println(sql);
                }
                System.out.println("打印多余的字段：" + rows.size());
//                赋值表名
                tableName = row.getCell(1).getStringCellValue().replace(" ", "");
                System.out.println("------------------------" + tableName + "------------------------");
                rows = new ArrayList<Row>();
                flag = true;
                continue;
            }
//            flag=true 说明这个是sql的，将其添加到rows，否则是文档的
            if (flag) {
                rows.add(row);
            } else {

//                此后每一行都和rows的数据进行对比看是不是一致的
                Iterator<Row> iterator = rows.iterator();
//                把这两个值赋值一样
                rowNumberBefore = rows.size();
                rowNumberAfter = rowNumberBefore;
//                当前行的字段
                String rowNewCellName = row.getCell(1).getStringCellValue().replace(" ", "").replace("\t", "");
                while (iterator.hasNext()) {
                    Row next = iterator.next();
                    String newCell = next.getCell(0).getStringCellValue().replace(" ", "")
                            .replace("[", "").replace("]", "");
//                    比较字段名是否相等
                    if (newCell.equals(rowNewCellName)) {

                        ++nameCount;
//                      判断数据类型是否相等
                        classSql = checkClass(row, next);

//                      判断是否为空
                        nullSql = checkNull(row, next);

//                        判断注释(不要注释)
//                        if (next.getCell(5).getStringCellValue() == "") {
//                            comment = " " + row.getCell(1).getStringCellValue();
//
//                        }
                        if (classSql == "" && nullSql != "") {
//                            System.out.println("打印这种情况：" + row.getCell(0).getStringCellValue().replace(" ", ""));
                            classSql = " " + row.getCell(3).getStringCellValue().replace(" ", "");
                        }

                        if (!(classSql + nullSql).equals("")) {
                            sql = "alter table [dbo].[" + tableName + "]  alter column " + newCell +
                                    classSql + nullSql;
//                            可能是否为空那项不存在 所以要判断
                            if (next.getCell(4) != null) {
                                if ("".equals(row.getCell(4).getStringCellValue().replace(" ", ""))) {
                                    nullFlag = "NOT NULL";
                                } else {
                                    nullFlag = "NULL";
                                }
                                if (next.getCell(4).getStringCellValue().replace(" ", "").
                                        replace(",", "").contains("NOT")) {
                                    nullOracleFlag = " not null";
                                } else {
                                    nullOracleFlag = "null";
                                }
                                System.out.println("字段：" + rowNewCellName + " sql类型：" + next.getCell(1).getStringCellValue().replace(" ", "") +
                                        " 文档类型：" + row.getCell(3).getStringCellValue().replace(" ", "") +
                                        " sql空值：" + nullOracleFlag + " 文档空值：" + nullFlag);
//                                System.out.println("这是需要修改的sql：");
                            }
                            System.out.println(sql + "\n");
                        }
//                        匹配到名字相等的就跳出循环
                        iterator.remove();
                        rowNumberAfter = rows.size();
                        break;
                    }

                }
//                没有  需要新增
                if (rowNumberBefore == rowNumberAfter) {
                    if ("INTERFACE_DETAIL".equals(tableName)) {
                        System.out.println();
                    }
//                    checkCellNull(row, 4);
                    System.out.println("没有这个字段：" + rowNewCellName + "\n");
                    if ("Y".equals(row.getCell(4).getStringCellValue().replace(" ", ""))) {
                        nullSql = " null";
                    }

                    if ("".equals(row.getCell(4).getStringCellValue().replace(" ", ""))) {
                        nullSql = " not null";
                    }

                    sql = "alert table " + tableName + " add " + rowNewCellName + " " +
                            row.getCell(3).getStringCellValue() + " " + nullSql + " comment '" +
                            row.getCell(1).getStringCellValue() + "' ;";
                    System.out.println("打印需要新增的的sql：");
                    System.out.println(sql);
                }
            }
        }
        System.out.println("完成！！！");
        System.out.println("打印类型不相等次数：" + count);
        System.out.println("打印空相等的次数：" + nullCount);
        System.out.println("打印名字相等的次数：" + nameCount);
        wb.close();
        input.close();
    }

    //    判断是否为空
    public static String checkNull(Row row, Row next) {
        String nullSql = "";
        if ("Y".equals(row.getCell(4).getStringCellValue().replace(" ", ""))) {
            nullSql = " null";
        }
        if ("".equals(row.getCell(4).getStringCellValue().replace(" ", ""))) {
            nullSql = " not null";
        }
        if ("".equals(row.getCell(4).getStringCellValue().replace(" ", ""))
                && "NOT".equals(next.getCell(4).getStringCellValue().replace(" ", ""))) {
            ++nullCount;
            nullSql = "";
            return nullSql;
        }

//        checkCellNull(next,2);
        if ("Y".equals(row.getCell(4).getStringCellValue().replace(" ", ""))
                && "NULL".equals(next.getCell(4).getStringCellValue()
                .replace(" ", "").replace(",", ""))) {
            ++nullCount;
            nullSql = "";
            return nullSql;
        }
//        System.out.println("打印空值不等的情况：" + row.getCell(0).getStringCellValue());
//        System.out.println("打印nullSql：" + nullSql);
        return nullSql;
    }

    //    判断类型是否对应
    public static String checkClass(Row row, Row next) {
        String classSql = "";
        if (!row.getCell(3).getStringCellValue().replace(" ", "").equals(next.getCell(1).getStringCellValue().replace(" ", ""))) {
            classSql = " " + row.getCell(3).getStringCellValue();
            ++count;
        }
        return classSql;
    }

    public static void checkCellNull(Row row, int i) {
        if (row.getCell(i) == null) {
            Cell cell = row.createCell(i);
            cell.setCellValue("");
        }
    }
}
