package com.example.xiaojin20135.excellib;

import java.util.ArrayList;
import java.util.Map;

public enum ExcelBean {
    EXCEL_BEAN;
    //每行 数据位每一列的title
    private ArrayList<String> titleList = new ArrayList<>();
    //填充的具体数据 每行数据为一个map,key为title，value为具体的数值
    private ArrayList<Map> rowList = new ArrayList<>();


    /*
    * @author lixiaojin
    * create on 2019/4/11 17:04
    * description: 清理当前数据
    */
    public void clearCurrent(){
        titleList.clear();
        rowList.clear();
    }

    /*
    * @author lixiaojin
    * create on 2019/4/11 17:05
    * description: 添加某列
    */
    public void addItem(String titleValue){
        titleList.add(titleValue);
    }

    /*
    * @author lixiaojin
    * create on 2019/4/11 17:07
    * description: 添加某行数据
    */
    public void addRowData(Map rowMap){
        rowList.add(rowMap);
    }

    /*
    * @author lixiaojin
    * create on 2019/4/11 17:55
    * description: 填充数据
    */
    public void initData(String[] titleList,ArrayList<Map> rowList){
        this.titleList.clear();
        for(int i=0;i<titleList.length;i++){
            this.titleList.add(titleList[i]);
        }
        this.rowList = rowList;
    }


    public ArrayList<String> getTitleList() {
        return titleList;
    }

    public ArrayList<Map> getRowList() {
        return rowList;
    }
}
