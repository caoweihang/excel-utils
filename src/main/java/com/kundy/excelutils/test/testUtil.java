package com.kundy.excelutils.test;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.util.CollectionUtils;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * @Auther: caoweihang
 * @Date: 2020/9/15 19:11
 * @Description:
 */
public class testUtil {


    public static void main(String[] args) throws Exception {
        //tring str = "\\/metrics$";

        List<List<String>> lists = readExcle("/Users/caoweihang/Desktop/rucang.xlsx");
        if (CollectionUtils.isEmpty(lists)){
            return;
        }
        StringBuilder builder = new StringBuilder();
        builder.append("insert into dffl_stock.goods_stock_enter_warehouse_migrate ")
                .append("(merchant_code,sku_code,new_merchant_code,merchant_sku_pk,new_merchant_sku_pk,`status`) values ");
        for (List<String> list : lists) {
            builder.append("\n").append("(")
                    .append(" '").append(list.get(2)).append("',")
                    .append(" '").append(list.get(3)).append("',")
                    .append(" '").append(list.get(4)).append("',")
                    .append(" '").append(list.get(1)).append("',")
                    .append(" '").append(list.get(5)).append("',")
                    .append(" 0 ),");
        }
        String result = builder.toString();

        String sql = result.substring(0, result.length() - 1);

        System.out.println(sql + ";");
    }

    public static List<List<String>> readExcle(String fileName) throws Exception {

        //new一个输入流
        FileInputStream inputStream = new FileInputStream(fileName);
        //new一个workbook
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
        //创建一个sheet对象，参数为sheet的索引
        XSSFSheet sheet = workbook.getSheetAt(0);
        //new出存放一张表的二维数组
        List<List<String>> allData = new ArrayList<List<String>>();

        for (Row row:sheet) {
            List<String> oneRow = new ArrayList<String>();
            //不读表头
            if(row.getRowNum()==0)
                continue;
            for (Cell cell : row) {
                cell.setCellType(CellType.STRING);
                oneRow.add(cell.getStringCellValue().trim());
            }
            allData.add(oneRow);
        }
        //关闭workbook
        workbook.close();
        return allData;
    }
}
