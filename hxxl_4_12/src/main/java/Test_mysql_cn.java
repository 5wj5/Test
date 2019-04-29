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

public class Test_mysql_cn {

    public static int count = 0;
    public static int nullCount = 0;
    public static int nameCount = 0;
    public static String tableName = null;

    public static void main(String[] args) throws IOException {
        InputStream input = Test_mysql_cn.class.getClassLoader().getResourceAsStream("123.xlsx");
        Workbook wb = new XSSFWorkbook(input);
        Sheet sheet = wb.getSheet("Sheet6");
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
                    if (tableName.equals("INTERFACE_MEDICAL_MATERIAL")) {
                        System.out.println();
                    }
                    sql = "alert table " + tableName + " drop " + value + ";";
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
                String rowNewCellName = row.getCell(1).getStringCellValue().replace(" ", "").replace("\t", "");

                while (iterator.hasNext()) {
                    Row next = iterator.next();
                    String newCell = next.getCell(0).getStringCellValue().replace(" ", "").replace("`", "");
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
                            sql = "alter table " + tableName + "  modify column " + newCell +
                                    classSql + nullSql + ";";
//                            可能是否为空那项不存在 所以要判断
                            if (next.getCell(2) != null) {
                                if ("".equals(row.getCell(4).getStringCellValue().replace(" ", ""))) {
                                    nullFlag = "NOT NULL";
                                } else {
                                    nullFlag = "DEFAULT NULL";
                                }
                                System.out.println("字段：" + rowNewCellName + " sql类型：" + next.getCell(1).getStringCellValue().replace(" ", "") +
                                        " 文档类型：" + row.getCell(3).getStringCellValue().replace(" ", "") +
                                        " sql空值：" + next.getCell(2).getStringCellValue().replace(" ", "") + " NULL" +
                                        " 文档空值：" + nullFlag);
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
                    checkCellNull(row, 4);
                    System.out.println("没有这个字段：" + rowNewCellName + "\n");
                    if ("Y".equals(row.getCell(4).getStringCellValue().replace(" ", ""))) {
                        nullSql = " default null";
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
            nullSql = " default null";
        }
        if ("".equals(row.getCell(4).getStringCellValue().replace(" ", ""))) {
            nullSql = " not null";
        }
        if ("INTERFACE_CASE".equals(tableName)) {
            System.out.println();
        }
        if (next.getCell(2) == null) {
            System.out.println("这是空");
        }
        if (next.getCell(2) != null && "".equals(row.getCell(4).getStringCellValue().replace(" ", ""))
                && "NOT".equals(next.getCell(2).getStringCellValue().replace(" ", ""))) {
            ++nullCount;
            nullSql = "";
            return nullSql;
        }

        checkCellNull(next,2);
        if ("Y".equals(row.getCell(4).getStringCellValue().replace(" ", ""))
                && "DEFAULT".equals(next.getCell(2).getStringCellValue().replace(" ", ""))) {
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
