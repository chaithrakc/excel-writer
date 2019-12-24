package com.yodlee.excelwriter.excelwriter;

import com.yodlee.excelwriter.exceptions.ExcelWriterException;
import com.yodlee.excelwriter.utils.ExcelHelper;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;

public class ComposeExcel<T> {
    private final static Logger LOGGER = Logger.getLogger(ComposeExcel.class);
    private SXSSFWorkbook workbook;
    private SXSSFSheet sheet;
    private File excelFile;
    private Class<T> excelVo;
    private Map<String,Integer> headerRowMapper = new HashMap<>();
    private AtomicInteger rowNum = new AtomicInteger();
    private ComposeExcel(File excelFile, Class<T> excelVoClass) throws IOException {
        isWriteableFile(excelFile);
        this.excelVo = excelVoClass;
        workbook = new SXSSFWorkbook(SXSSFWorkbook.DEFAULT_WINDOW_SIZE);
        sheet = workbook.createSheet(this.excelVo.getSimpleName());
    }
    public ComposeExcel(String excelFilePath, Class<T> excelVoClass) throws IOException{
        this(new File(excelFilePath), excelVoClass);
    }
    public void renderData(List<T> tList) throws ExcelWriterException {
        getExcelVoFields();
        createHeaderRow();
        generateExcelRows(tList);
    }
    private void generateExcelRows(List<T> tList) throws ExcelWriterException {
        for (T vo : tList) {
            Row row = sheet.createRow(rowNum.incrementAndGet());
            try {
                populateExcelRows(row, vo);
            } catch (ExcelWriterException e) {
                e.printStackTrace();
                throw e;
            }
        }
        LOGGER.info("rows written "+rowNum);
    }
    private void getExcelVoFields() {
        Field[] declaredFields = excelVo.getDeclaredFields();
        AtomicInteger columnNum = new AtomicInteger(0);
        Arrays.stream(declaredFields).forEach(field -> headerRowMapper.put(ExcelHelper.capitalizeInitialLetter(field.getName()), columnNum.getAndIncrement()));
    }
    private void createHeaderRow() {
        Row headerRow = sheet.createRow(0);
        headerRowMapper.forEach((headerName, columnNum) ->
                formatHeaderRow(createCell(headerRow, columnNum, headerName))
        );
    }
    private void formatHeaderRow(Cell headerCell) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setColor(IndexedColors.BLACK.getIndex());
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerCell.setCellStyle(style);
    }
    public void generateExcelFile() throws ExcelWriterException, IOException {
        FileOutputStream fileOutputStream = null;
        try {
            fileOutputStream = new FileOutputStream(excelFile);
            workbook.write(fileOutputStream);
            fileOutputStream.close();
        } catch (IOException ex) {
            throw new ExcelWriterException("unable to generate excel file " + ex.getMessage());
        } finally {
            if (fileOutputStream != null)
                fileOutputStream.close();
            workbook.close();
        }
    }
    private void populateExcelRows(Row row, T t) throws ExcelWriterException {
        AtomicReference<Object> cellValue = new AtomicReference<>();
        String fieldName;
        Integer columnNum;
        for (Map.Entry<String, Integer> entry : headerRowMapper.entrySet()) {
            fieldName = entry.getKey();
            columnNum = entry.getValue();
            try {
                Method method = t.getClass().getMethod("get".concat(fieldName));
                cellValue.set(method.invoke(t));
                createCell(row, columnNum, cellValue.get());
            } catch (Exception ex) {
                throw new ExcelWriterException("unable to write data for columnNum" + columnNum + " fieldName " + fieldName);
            }
        }
    }
    private void isWriteableFile(File excelFile) throws IOException {
        if (!excelFile.canWrite())
            throw new IOException("Unable to write " + excelFile.getAbsolutePath());
        this.excelFile = excelFile;
    }
    private Cell createCell(Row row, Integer columnNumber, Object cellValue) {
        Cell cell = row.createCell(columnNumber);
        cell.setCellValue(cellValue != null ? cellValue.toString().trim() : ExcelHelper.STRING_EMPTY);
        return cell;
    }
}