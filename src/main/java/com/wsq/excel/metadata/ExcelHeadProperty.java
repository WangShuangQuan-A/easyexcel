package com.wsq.excel.metadata;

import com.wsq.excel.annotation.ExcelColumnNum;
import com.wsq.excel.annotation.ExcelProperty;

import java.lang.reflect.Field;
import java.util.*;

/**
 * Define the header attribute of excel
 *
 * @author jipengfei
 */
public class ExcelHeadProperty {

    /**
     * Custom class
     */
    private Class<? extends BaseRowModel> headClazz;

    /**
     * A two-dimensional array describing the header
     */
    private List<String> head;

    /**
     * Attributes described by the header
     */
    private List<ExcelColumnProperty> columnPropertyList = new ArrayList<>();

    /**
     * Attributes described by the header
     */
    private Map<Integer, ExcelColumnProperty> excelColumnPropertyMap1 = new HashMap<>();
    private Map<String, ExcelColumnProperty> excelColumnPropertyMap2 = new HashMap<>();

    public ExcelHeadProperty(Class<? extends BaseRowModel> headClazz, List<String> head) {
        this.headClazz = headClazz;
        this.head = head;
        initColumnProperties();
    }

    /**
     */
    private void initColumnProperties() {
        if (this.headClazz != null) {
            List<Field> fieldList = new ArrayList<Field>();
            Class tempClass = this.headClazz;
            //When the parent class is null, it indicates that the parent class (Object class) has reached the top
            // level.
            while (tempClass != null) {
                fieldList.addAll(Arrays.asList(tempClass.getDeclaredFields()));
                //Get the parent class and give it to yourself
                tempClass = tempClass.getSuperclass();
            }
            //List<List<String>> headList = new ArrayList<List<String>>();
            for (Field f : fieldList) {
                initOneColumnProperty(f);
            }
            //对列排序
            Collections.sort(columnPropertyList);
            if (head == null || head.size() == 0) {
                head = new ArrayList<>();
                for (ExcelColumnProperty excelColumnProperty : columnPropertyList) {
                    head.add(excelColumnProperty.getHead());
                }
            }
        }
    }

    /**
     * @param f
     */
    private void initOneColumnProperty(Field f) {
        ExcelProperty p = f.getAnnotation(ExcelProperty.class);
        ExcelColumnProperty excelHeadProperty = null;
        if (p != null) {
            excelHeadProperty = new ExcelColumnProperty();
            excelHeadProperty.setField(f);
            excelHeadProperty.setHead(p.value());
            excelHeadProperty.setIndex(p.index());
            excelHeadProperty.setFormat(p.format());
            if (p.index() == ExcelColumnProperty.DEFAULT_INDEX) {
                excelColumnPropertyMap2.put(p.value(), excelHeadProperty);
            }else {
                excelColumnPropertyMap1.put(p.index(), excelHeadProperty);
            }
        } else {
            ExcelColumnNum columnNum = f.getAnnotation(ExcelColumnNum.class);
            if (columnNum != null) {
                excelHeadProperty = new ExcelColumnProperty();
                excelHeadProperty.setField(f);
                excelHeadProperty.setIndex(columnNum.value());
                excelHeadProperty.setFormat(columnNum.format());
                excelColumnPropertyMap1.put(columnNum.value(), excelHeadProperty);
            }
        }
        if (excelHeadProperty != null) {
            this.columnPropertyList.add(excelHeadProperty);
        }

    }

    /**
     * 一行数据
     */
    public void appendOneRow(List<String> row) {

        for (int i = 0; i < row.size(); i++) {
            String str;
            if (head.size() <= i) {
                head.add("");
            } /*else {
                list = head.get(0);
            }
            list.add(row.get(i));*/
        }

    }

    /**
     * @param columnNum
     * @return
     */
    public ExcelColumnProperty getExcelColumnProperty(int columnNum) {
        return excelColumnPropertyMap1.get(columnNum);
    }

    public Class getHeadClazz() {
        return headClazz;
    }

    public void setHeadClazz(Class headClazz) {
        this.headClazz = headClazz;
    }

    public List<String> getHead() {
        return this.head;
    }

    public void setHead(List<String> head) {
        this.head = head;
    }
    public List<ExcelColumnProperty> getColumnPropertyList() {
        return columnPropertyList;
    }

    public void setColumnPropertyList(List<ExcelColumnProperty> columnPropertyList) {
        this.columnPropertyList = columnPropertyList;
    }

    /**
     * Calculate all cells that need to be merged
     *
     * @return cells that need to be merged
     */
    public List<CellRange> getCellRangeModels() {
        List<CellRange> cellRanges = new ArrayList<>();
   /*     for (int i = 0; i < head.size(); i++) {
            String columnValue = head.get(i);
            for (int j = 0; j < columnValues.size(); j++) {
                int lastRow = getLastRangNum(j, columnValues.get(j), columnValues);
                int lastColumn = getLastRangNum(i, columnValues.get(j), getHeadByRowNum(j));
                if ((lastRow > j || lastColumn > i) && lastRow >= 0 && lastColumn >= 0) {
                    cellRanges.add(new CellRange(j, lastRow, i, lastColumn));
                }
            }
        }*/
        return cellRanges;
    }

    public String getHeadByRowNum(int rowNum) {
        if (head.size() > rowNum) {
            return head.get(rowNum);
        }
        return "noHead";
    }

    /**
     * Get the last consecutive string position
     *
     * @param j      current value position
     * @param value  value content
     * @param values values
     * @return the last consecutive string position
     */
    private int getLastRangNum(int j, String value, List<String> values) {
        if (value == null) {
            return -1;
        }
        if (j > 0) {
            String preValue = values.get(j - 1);
            if (value.equals(preValue)) {
                return -1;
            }
        }
        int last = j;
        for (int i = last + 1; i < values.size(); i++) {
            String current = values.get(i);
            if (value.equals(current)) {
                last = i;
            } else {
                // if i>j && !value.equals(current) Indicates that the continuous range is exceeded
                if (i > j) {
                    break;
                }
            }
        }
        return last;

    }

    public int getRowNum() {
        int headRowNum = 0;
        if (head != null && head.size() > 0) {
            headRowNum = head.size();
        }
        return headRowNum;
    }

    public Map<Integer, ExcelColumnProperty> getExcelColumnPropertyMap1() {
        return excelColumnPropertyMap1;
    }

    public void setExcelColumnPropertyMap1(Map<Integer, ExcelColumnProperty> excelColumnPropertyMap1) {
        this.excelColumnPropertyMap1 = excelColumnPropertyMap1;
    }

    public Map<String, ExcelColumnProperty> getExcelColumnPropertyMap2() {
        return excelColumnPropertyMap2;
    }

    public void setExcelColumnPropertyMap2(Map<String, ExcelColumnProperty> excelColumnPropertyMap2) {
        this.excelColumnPropertyMap2 = excelColumnPropertyMap2;
    }
}
