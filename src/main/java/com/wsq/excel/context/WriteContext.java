package com.wsq.excel.context;

import com.wsq.excel.event.WriteHandler;
import com.wsq.excel.metadata.BaseRowModel;
import com.wsq.excel.metadata.ExcelHeadProperty;
import com.wsq.excel.metadata.Table;
import com.wsq.excel.support.ExcelTypeEnum;
import com.wsq.excel.util.CollectionUtils;
import com.wsq.excel.util.StyleUtil;
import com.wsq.excel.util.WorkBookUtil;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

import static com.wsq.excel.util.StyleUtil.buildSheetStyle;

/**
 * A context is the main anchorage point of a excel writer.
 *
 * @author jipengfei
 */
public class WriteContext {

    /***
     * The sheet currently written
     */
    private Sheet currentSheet;

    /**
     * current param
     */
    private com.wsq.excel.metadata.Sheet currentSheetParam;

    /**
     * The sheet currently written's name
     */
    private String currentSheetName;

    /**
     *
     */
    private Table currentTable;

    /**
     * Excel type
     */
    private ExcelTypeEnum excelType;

    /**
     * POI Workbook
     */
    private Workbook workbook;

    /**
     * Final output stream
     */
    private OutputStream outputStream;

    /**
     * Written form collection
     */
    private Map<Integer, Table> tableMap = new ConcurrentHashMap<Integer, Table>();

    /**
     * Cell default style
     */
    private CellStyle defaultCellStyle;

    /**
     * Current table head  style
     */
    private CellStyle currentHeadCellStyle;

    /**
     * Current table content  style
     */
    private CellStyle currentContentCellStyle;

    /**
     * the header attribute of excel
     */
    private ExcelHeadProperty excelHeadProperty;

    private boolean needHead = Boolean.TRUE;

    private WriteHandler afterWriteHandler;

    public WriteHandler getAfterWriteHandler() {
        return afterWriteHandler;
    }

    public WriteContext(InputStream templateInputStream, OutputStream out, ExcelTypeEnum excelType,
                        boolean needHead, WriteHandler afterWriteHandler) throws IOException {
        this.needHead = needHead;
        this.outputStream = out;
        this.afterWriteHandler = afterWriteHandler;
        this.workbook = WorkBookUtil.createWorkBook(templateInputStream, excelType);
        this.defaultCellStyle = StyleUtil.buildDefaultCellStyle(this.workbook);

    }

    /**
     * @param sheet
     */
    public void currentSheet(com.wsq.excel.metadata.Sheet sheet) {
        if (null == currentSheetParam || currentSheetParam.getSheetNo() != sheet.getSheetNo()) {
            cleanCurrentSheet();
            currentSheetParam = sheet;
            try {
                this.currentSheet = workbook.getSheetAt(sheet.getSheetNo() - 1);
            } catch (Exception e) {
                this.currentSheet = WorkBookUtil.createSheet(workbook, sheet);
                if (null != afterWriteHandler) {
                    this.afterWriteHandler.sheet(sheet.getSheetNo(), currentSheet);
                }
            }
            buildSheetStyle(currentSheet, sheet.getColumnWidthMap());
            /** **/
            this.initCurrentSheet(sheet);
        }

    }

    private void initCurrentSheet(com.wsq.excel.metadata.Sheet sheet) {

        /** **/
        initExcelHeadProperty(sheet.getHead(), sheet.getClazz());

        initTableStyle(sheet.getTableStyle());

        initTableHead();

    }

    private void cleanCurrentSheet() {
        this.currentSheet = null;
        this.currentSheetParam = null;
        this.excelHeadProperty = null;
        this.currentHeadCellStyle = null;
        this.currentContentCellStyle = null;
        this.currentTable = null;

    }

    /**
     * init excel header
     *
     * @param head
     * @param clazz
     */
    private void initExcelHeadProperty(List<String> head, Class<? extends BaseRowModel> clazz) {
        if (head != null || clazz != null) { this.excelHeadProperty = new ExcelHeadProperty(clazz, head); }
    }

    public void initTableHead() {
        if (needHead && null != excelHeadProperty && !CollectionUtils.isEmpty(excelHeadProperty.getHead())) {
            int startRow = currentSheet.getLastRowNum();
            if (startRow > 0) {
                startRow += 4;
            } else {
                startRow = currentSheetParam.getStartRow();
            }
            addMergedRegionToCurrentSheet(startRow);
            int i = startRow;
            for (; i < this.excelHeadProperty.getRowNum() + startRow; i++) {
                Row row = WorkBookUtil.createRow(currentSheet, i);
                if (null != afterWriteHandler) {
                    this.afterWriteHandler.row(i, row);
                }
                addOneRowOfHeadDataToExcel(row, this.excelHeadProperty.getHeadByRowNum(i - startRow));
            }
        }
    }

    private void addMergedRegionToCurrentSheet(int startRow) {
        for (com.wsq.excel.metadata.CellRange cellRangeModel : excelHeadProperty.getCellRangeModels()) {
            currentSheet.addMergedRegion(new CellRangeAddress(cellRangeModel.getFirstRow() + startRow,
                cellRangeModel.getLastRow() + startRow,
                cellRangeModel.getFirstCol(), cellRangeModel.getLastCol()));
        }
    }

    private void addOneRowOfHeadDataToExcel(Row row,String headByRowNum) {
        if (headByRowNum != null) {
            Cell cell = WorkBookUtil.createCell(row, 0, getCurrentHeadCellStyle(), headByRowNum);
            if (null != afterWriteHandler) {
                this.afterWriteHandler.cell(0, cell);
            }
        }
    }

    private void initTableStyle(com.wsq.excel.metadata.TableStyle tableStyle) {
        if (tableStyle != null) {
            this.currentHeadCellStyle = StyleUtil.buildCellStyle(this.workbook, tableStyle.getTableHeadFont(),
                tableStyle.getTableHeadBackGroundColor());
            this.currentContentCellStyle = StyleUtil.buildCellStyle(this.workbook, tableStyle.getTableContentFont(),
                tableStyle.getTableContentBackGroundColor());
        }
    }

    private void cleanCurrentTable() {
        this.excelHeadProperty = null;
        this.currentHeadCellStyle = null;
        this.currentContentCellStyle = null;
        this.currentTable = null;

    }

    public void currentTable(Table table) {
        if (null == currentTable || currentTable.getTableNo() != table.getTableNo()) {
            cleanCurrentTable();
            this.currentTable = table;
            this.initExcelHeadProperty(table.getHead(), table.getClazz());
            this.initTableStyle(table.getTableStyle());
            this.initTableHead();
        }

    }

    public ExcelHeadProperty getExcelHeadProperty() {
        return this.excelHeadProperty;
    }

    public boolean needHead() {
        return this.needHead;
    }

    public Sheet getCurrentSheet() {
        return currentSheet;
    }

    public void setCurrentSheet(Sheet currentSheet) {
        this.currentSheet = currentSheet;
    }

    public String getCurrentSheetName() {
        return currentSheetName;
    }

    public void setCurrentSheetName(String currentSheetName) {
        this.currentSheetName = currentSheetName;
    }

    public ExcelTypeEnum getExcelType() {
        return excelType;
    }

    public void setExcelType(ExcelTypeEnum excelType) {
        this.excelType = excelType;
    }

    public OutputStream getOutputStream() {
        return outputStream;
    }

    public CellStyle getCurrentHeadCellStyle() {
        return this.currentHeadCellStyle == null ? defaultCellStyle : this.currentHeadCellStyle;
    }

    public CellStyle getCurrentContentStyle() {
        return this.currentContentCellStyle;
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    public com.wsq.excel.metadata.Sheet getCurrentSheetParam() {
        return currentSheetParam;
    }

    public void setCurrentSheetParam(com.wsq.excel.metadata.Sheet currentSheetParam) {
        this.currentSheetParam = currentSheetParam;
    }

    public Table getCurrentTable() {
        return currentTable;
    }

    public void setCurrentTable(Table currentTable) {
        this.currentTable = currentTable;
    }
}


