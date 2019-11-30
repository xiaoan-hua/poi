package com.baizhi;

import com.baizhi.entity.User;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

@SpringBootTest
class PoiApplicationTests {

    @Test
    void contextLoads() throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet();
        HSSFRow row = sheet.createRow(0);
        row.createCell(0).setCellValue("id");
        row.createCell(1).setCellValue("name");
        row.createCell(2).setCellValue("age");
        row.createCell(3).setCellValue("bir");
        FileOutputStream outputStream = null;
        try {
            outputStream = new FileOutputStream("E:/fzz.xls");
            workbook.write(outputStream);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            outputStream.close();
        }

    }

    @Test
    void contextLoads1() throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet();
        //设置表中某列的宽度
        sheet.setColumnWidth(3, 15 * 256);
        //设置数据的时间格式
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        short format = workbook.createDataFormat().getFormat("yyyy-MM-dd");
        cellStyle.setDataFormat(format);
        //设置表头的样式
        HSSFCellStyle cellStyle1 = workbook.createCellStyle();
        HSSFFont font = workbook.createFont();
        font.setBold(true);
        font.setColor(Font.COLOR_RED);
        font.setItalic(true);
        font.setFontName("仿宋");
        cellStyle1.setFont(font);
        cellStyle1.setAlignment(HorizontalAlignment.CENTER);
        //创建表
        HSSFRow row = sheet.createRow(0);
        HSSFCell cell1 = row.createCell(0);
        cell1.setCellStyle(cellStyle1);
        cell1.setCellValue("id");
        row.createCell(1).setCellValue("name");
        row.createCell(2).setCellValue("age");
        row.createCell(3).setCellValue("bir");
        List<User> list = new ArrayList<>();
        User user = new User("1", "13", "fzz", new Date());
        User user1 = new User("2", "15", "duidui", new Date());
        User user2 = new User("3", "14", "zz", new Date());
        list.add(user);
        list.add(user1);
        list.add(user2);
        for (int i = 0; i < list.size(); i++) {
            HSSFRow row1 = sheet.createRow(i + 1);
            row1.createCell(0).setCellValue(list.get(i).getId());
            row1.createCell(1).setCellValue(list.get(i).getAge());
            row1.createCell(2).setCellValue(list.get(i).getName());
            HSSFCell cell = row1.createCell(3);
            cell.setCellStyle(cellStyle);
            cell.setCellValue(list.get(i).getBir());
        }

        FileOutputStream outputStream = null;
        try {
            outputStream = new FileOutputStream("E:/fzz.xls");
            workbook.write(outputStream);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            outputStream.close();
        }

    }

}
