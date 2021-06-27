package com.reports.excelwriter;

import com.reports.exceptions.ExcelWriterException;
import com.reports.utils.ExcelHelper;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;

public class ExcelWriter<T> {
    private static final Logger LOGGER = LogManager.getLogger(ExcelWriter.class);
    private final SXSSFWorkbook workbook;
    private final SXSSFSheet sheet;
    private File excelFile;
    private final Class<T> excelVo;
    private final Map<String,Integer> headerRowMapper = new HashMap<>();
    private final AtomicInteger rowNum = new AtomicInteger();
    private ExcelWriter(File excelFile, Class<T> excelVoClass) {
        isWriteableFile(excelFile);
        this.excelVo = excelVoClass;
        workbook = new SXSSFWorkbook(SXSSFWorkbook.DEFAULT_WINDOW_SIZE);
        sheet = workbook.createSheet(this.excelVo.getSimpleName());
    }
    public ExcelWriter(String excelFilePath, Class<T> excelVoClass) {
        this(new File(excelFilePath), excelVoClass);
    }
    public void renderData(List<T> excelVoList) throws ExcelWriterException {
        getExcelVoFields();
        createHeaderRow();
        generateExcelRows(excelVoList);
    }
    private void generateExcelRows(List<T> excelVoList) throws ExcelWriterException {
        for (T vo : excelVoList) {
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
    public void generateExcelFile() throws ExcelWriterException {
        try(FileOutputStream fileOutputStream = new FileOutputStream(excelFile)) {
            workbook.write(fileOutputStream);
        } catch (IOException ex) {
            throw new ExcelWriterException("unable to generate excel file " + ex.getMessage());
        }
    }
    private void populateExcelRows(Row row, T excelVo) throws ExcelWriterException {
        AtomicReference<Object> cellValue = new AtomicReference<>();
        String fieldName;
        Integer columnNum;
        for (Map.Entry<String, Integer> entry : headerRowMapper.entrySet()) {
            fieldName = entry.getKey();
            columnNum = entry.getValue();
            try {
                Method method = excelVo.getClass().getMethod("get".concat(fieldName));
                cellValue.set(method.invoke(excelVo));
                createCell(row, columnNum, cellValue.get());
            } catch (Exception ex) {
                throw new ExcelWriterException("unable to write data for columnNum" + columnNum + " fieldName " + fieldName);
            }
        }
    }
    private void isWriteableFile(File excelFile) {
        this.excelFile = excelFile;
    }
    private Cell createCell(Row row, Integer columnNumber, Object cellValue) {
        Cell cell = row.createCell(columnNumber);
        cell.setCellValue(cellValue != null ? cellValue.toString().trim() : ExcelHelper.STRING_EMPTY);
        return cell;
    }
}