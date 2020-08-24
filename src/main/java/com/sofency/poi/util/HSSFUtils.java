package com.sofency.poi.util;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFCreationHelper;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;

/**
 * @author xiaohufeng
 * @date:
 */
public class HSSFUtils {
    public static HSSFWorkbook workbook = new HSSFWorkbook();
    /**
     * TODO
     * 创建工作空间
     * @date 2020/8/24
     * @param name
     * @return {@link HSSFSheet}
     */
    public static HSSFSheet createSheet(String name){
        HSSFSheet sheet =workbook.createSheet(name);//创建工作空间
        return sheet;
    }

    /**
     * TODO
     *　日期的设置
     * @date 2020/8/24
     * @return {@link null}
     */
    public static  HSSFCellStyle createDateStyle(String format){
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        HSSFCreationHelper creationHelper = workbook.getCreationHelper();
        cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat(format));
        return cellStyle;
    }

    /**
     * TODO
     *　设置几位小数点
     * @date 2020/8/24
     * @return {@link null}
     */
    public static HSSFCellStyle createNumber(int number){
        StringBuilder format = new StringBuilder();
        if(number<0){
            throw new IllegalArgumentException("参数不合法");
        }else if(number==0){
            format.append("0");
        }else{
            format.append("0.");
            for(int i=1;i<=number;i++){
                format.append("0");
            }
        }
        HSSFCellStyle cellStyle1 = workbook.createCellStyle();
        cellStyle1.setDataFormat(HSSFDataFormat.getBuiltinFormat(format.toString()));
        return cellStyle1;
    }

    public static void close(OutputStream output) throws IOException {
        workbook.setActiveSheet(0);
        workbook.write(output);//写到输出流中
        output.close();
    }

}
