package com.sofency.poi.file;

import com.sofency.poi.util.HSSFUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFCreationHelper;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Date;
import javax.sound.midi.Soundbank;
import javax.swing.filechooser.FileSystemView;

/**
 * @author xiaohufeng
 * @date:
 */
public class CreateFile {

    public  void CreateFile(String path) throws IOException {

        File file = new File(path);//创建文件
        OutputStream output = new FileOutputStream(file);

        HSSFSheet sheet = HSSFUtils.createSheet("sheet");
        HSSFRow row  = sheet.createRow(0);

        row.createCell(0).setCellValue("id");
        row.createCell(1).setCellValue("订单号");
        row.createCell(2).setCellValue("下单时间");
        row.createCell(3).setCellValue("个数");
        row.createCell(4).setCellValue("单价");
        row.createCell(5).setCellValue("订单金额");
        row.setHeightInPoints(30);//设置行高

        HSSFRow row1 = sheet.createRow(1);
        row1.createCell(0).setCellValue("1");
        row1.createCell(1).setCellValue("NOOO1");

        //日期格式化  设置样式
        sheet.setColumnWidth(2,20*256);

        HSSFCell cell2 = row1.createCell(2);
        cell2.setCellStyle(HSSFUtils.createDateStyle("yyyy-MM-dd HH:mm:ss"));
        cell2.setCellValue(new Date());

        row1.createCell(3).setCellValue(9);



        HSSFCell cell4 = row1.createCell(4);
        cell4.setCellStyle(HSSFUtils.createNumber(2));
        cell4.setCellValue(21.2222);


        //货币格式化
        HSSFCellStyle cellStyle2 = HSSFUtils.workbook.createCellStyle();
        HSSFFont font = HSSFUtils.workbook.createFont();
        font.setFontName("华文行楷");
        font.setFontHeightInPoints((short)15);
        font.setColor(HSSFColor.RED.index);
        cellStyle2.setFont(font);
        cellStyle2.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00"));

        HSSFCell cell5 = row1.createCell(5);
        cell5.setCellStyle(cellStyle2);
        cell5.setCellFormula("D2*E2");//设置计算公式

        HSSFUtils.close(output);
    }

}
