package com.example.xiaojin20135.excelapp;

import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.view.View;
import android.widget.Button;

import com.example.xiaojin20135.excellib.ExcelBean;
import com.example.xiaojin20135.excellib.ExcelExportHelper;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

public class MainActivity extends AppCompatActivity {

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
    }

    public void export(View view) {
        String[] titleArr = new String[]{"第一列","第二列","第三列","第四列","第五列","第六列"};
        ArrayList<Map> rowList = new ArrayList<>();
        Map map1 = new HashMap();
        for(int i=0;i<titleArr.length;i++){
            map1.put(titleArr[i],titleArr[i] + "的数据");
        }
        rowList.add(map1);
        Map map2 = new HashMap();
        for(int i=0;i<titleArr.length;i++){
            map2.put(titleArr[i],titleArr[i] + "的数据2");
        }
        rowList.add(map2);
        Map map3 = new HashMap();
        for(int i=0;i<titleArr.length;i++){
            map3.put(titleArr[i],titleArr[i] + "的数据3");
        }
        rowList.add(map3);

        ExcelBean.EXCEL_BEAN.initData(titleArr,rowList);
        ExcelExportHelper.EXCEL_EXPORT_HELPER.exportToExcelCommon(this,"haha","filename");
    }
}
