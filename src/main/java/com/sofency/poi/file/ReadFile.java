package com.sofency.poi.file;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.BufferedInputStream;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.SortedMap;

/**
 * @author xiaohufeng
 * @date:
 */
public class ReadFile {

    public static void readFile(String path) throws IOException {
        FileInputStream input = new FileInputStream(path);
        BufferedInputStream bufferedInputStream = new BufferedInputStream(input);
        POIFSFileSystem fileSystem = new POIFSFileSystem(bufferedInputStream);
        HSSFWorkbook workbook = new HSSFWorkbook(fileSystem);
        HSSFSheet sheet = workbook.createSheet("sheet");
        int lastRow = sheet.getLastRowNum();
        System.out.println(lastRow);
        for(int i = 0; i<lastRow;i++){
            HSSFRow row = sheet.getRow(i);
            if(row==null) break;
            short lastCellNum = row.getLastCellNum();
            for(int j=0; j<lastCellNum;j++){
                String cellValue = row.getCell(j).getStringCellValue();
                System.out.println(cellValue);
            }
        }

    }


}
