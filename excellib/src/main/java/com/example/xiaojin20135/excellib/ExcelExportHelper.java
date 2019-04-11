package com.example.xiaojin20135.excellib;
import android.app.Activity;
import android.content.Context;
import android.content.DialogInterface;
import android.content.Intent;
import android.net.Uri;
import android.os.Environment;
import android.support.v7.app.AlertDialog;
import android.util.Log;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Map;

import schemasMicrosoftComOfficeOffice.STInsetMode;

/**
 * Created by xiaojin20135 on 2017-12-13.
 * 测量结果导出到excel帮助类
 */
public enum ExcelExportHelper {
    EXCEL_EXPORT_HELPER;
    private static final String TAG = "ExcelExportHelper";
    public boolean exportToExcelCommon(Context context, String directoryName, String fileName){
        //创建文件夹
        try {
            new AppExternalFileWriter(context).createAppDirectory(directoryName);
        } catch (AppExternalFileWriter.ExternalFileWriterException e) {
            e.printStackTrace();
        }
        try{
            String tag = ".xlsx";
            String excelName = fileName + tag;
            //行索引
            int rowIndex = 1;

            //创建一个工作簿
            XSSFWorkbook workbook = new XSSFWorkbook();
            //创建一个电子表格
            XSSFSheet sheet = workbook.createSheet("mySheet");
            //行对象
            XSSFRow row ;

            String path = Environment.getExternalStorageDirectory().getPath()+ "/"+ directoryName;
            File file = new File(path , excelName);
            //判断文件是否存在
            boolean isExist = file.exists();
            if(isExist){//如果已经存在，在原文件基础上添加
                workbook = getWorkBookHandler(directoryName,excelName);
                sheet = workbook.getSheet("mySheet");
                rowIndex = rowIndex + sheet.getLastRowNum();
            }
            //获取列标题数据
            ArrayList<String> titleList = ExcelBean.EXCEL_BEAN.getTitleList();
            //获取需要写入的数据
            ArrayList<Map> rowList = ExcelBean.EXCEL_BEAN.getRowList();
            if(titleList != null && rowList != null && titleList.size() > 0 && rowList.size() > 0){
                int datasLen = rowList.size();
                int titleSize = titleList.size();
                if(rowIndex == 1){
                    //生成第一行
                    row = sheet.createRow(0);
                    //填充第一行数据
                    for(int t=0;t<titleSize;t++){
                        Cell cell = row.createCell(t);
                        cell.setCellValue(titleList.get(t));
                    }
                }

                //从第二行开始填充具体数据
                for(int i=0;i<datasLen;i++){
                    //当前行数据map
                    Map rowMap = rowList.get(i);
                    //创建一个行对象
                    row = sheet.createRow(i+rowIndex);
                    for(int titleIndex = 0;titleIndex<titleSize;titleIndex++){
                        Cell cell = row.createCell(titleIndex);
                        cell.setCellValue(getValueByKey(rowMap,titleList.get(titleIndex)));
                    }
                }
            }

            FileOutputStream fos = new FileOutputStream(file);
            workbook.write(fos);
            workbook.close();
            scanFile(context,path,fileName,tag);
            return true;
        }catch (Exception e){
            e.printStackTrace();
            AlertDialog.Builder builder = new AlertDialog.Builder(context);
            builder.setMessage("导出失败！" + e.getMessage());
            builder.setCancelable(true);
            builder.setPositiveButton("确定", new DialogInterface.OnClickListener() {
                @Override
                public void onClick(DialogInterface dialog, int which) {
                    ;
                }
            });
            builder.show();
            return false;
        }
    }

    /*
    * @author lixiaojin
    * create on 2019/4/11 17:20
    * description: 生成文件之后扫描文件，生效
    */
    private void scanFile(Context context,String path,String fileName,String tag){
        String filePath = path + "/" + fileName+tag;
        Log.d(TAG,"filePath = " + filePath);
        Intent scanIntent = new Intent(Intent.ACTION_MEDIA_SCANNER_SCAN_FILE);
        scanIntent.setData(Uri.fromFile(new File(filePath)));
        context.sendBroadcast(scanIntent);
    }


    private XSSFWorkbook getWorkBookHandler(String debug_file_dir, String excelName){
        XSSFWorkbook wb = null;
        FileInputStream fileInputStream = null;
        File file = new File(Environment.getExternalStorageDirectory().getPath()+ "/"+ debug_file_dir , excelName);
        try{
            if(file != null){
                fileInputStream = new FileInputStream(file);
                wb = new XSSFWorkbook(fileInputStream);
            }
        }catch (Exception e){
            return null;
        }finally {
            if(fileInputStream != null){
                try {
                    fileInputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return wb;
    }


    private String getValueByKey(Map map,String key){
        if(map != null && map.get(key) != null){
            return map.get(key).toString();
        }else{
            return "";
        }
    }
}
