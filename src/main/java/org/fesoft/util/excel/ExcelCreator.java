package org.fesoft.util.excel;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.Comparator;
import java.util.Date;
import java.util.List;
import java.util.Locale;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public class ExcelCreator {

    Class clazz;
    List<Column> columns;

    public Workbook create(List<?> data) {
        clazz = data.get(0).getClass();
        columns = findExcelColumns();
        return createWorkbook(data);
    }

    private Workbook createWorkbook(List<?> data) {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("sheet1");
        sheet.setDisplayGridlines(false);
        createHeader(sheet);
        createBody(sheet, data);
        return wb;
    }

    private List<Column> findExcelColumns() {
        List<Column> excelColumns = Stream.of(clazz.getDeclaredFields())
                .filter(field -> field.getAnnotation(ExcelColumn.class) != null)
                .map(f -> {
                    ExcelColumn excelColumn = f.getAnnotation(ExcelColumn.class);
                    return new Column(excelColumn.number(), excelColumn.header(), f.getName(), createMethod(f));
                }).collect(Collectors.toList());
        return excelColumns.stream().sorted(Comparator.comparing(Column::getNumber)).collect(Collectors.toList());
    }

    private Method createMethod(Field field) {
        Method method = null;
        String methodName = "get" + field.getName().substring(0, 1).toUpperCase(Locale.ENGLISH)
                + field.getName().substring(1);
        try {
            method = clazz.getMethod(methodName);
        } catch (NoSuchMethodException e) {
            e.printStackTrace();
        }
        return method;
    }


    private void createHeader(Sheet sheet) {
        CellStyle headerStyle = createHeaderStyle(sheet);
        List<String> headers = columns.stream().map(Column::getHeader).collect(Collectors.toList());
        Row row = sheet.createRow(0);
        for (int i = 0; i < headers.size(); i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(headers.get(i));
            cell.setCellStyle(headerStyle);
        }
    }

    private CellStyle createHeaderStyle(Sheet sheet) {
        CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        cellStyle.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        XSSFFont fontHeader = (XSSFFont) sheet.getWorkbook().createFont();
        fontHeader.setBold(true);
        cellStyle.setFont(fontHeader);
        return cellStyle;
    }

    private void createBody(Sheet sheet, List<?> data) {
        CellStyle cellStyle = createBodyCellStyle(sheet);
        for (int i = 0; i < data.size(); i++) {
            Row row = sheet.createRow(i + 1);
            Object rowData = data.get(i);
            writeRowData(rowData, row, cellStyle);
        }

    }

    private CellStyle createBodyCellStyle(Sheet sheet) {
        CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        return cellStyle;
    }

    private void writeRowData(Object rowData, Row row, CellStyle cellStyle) {
        for (int i = 0; i < columns.size(); i++) {
            Column column = columns.get(i);
            Cell cell = row.createCell(i);
            cell.setCellStyle(cellStyle);
            try {
                Object cellData = column.getMethodToSet().invoke(rowData);
                writeToCell(cell, cellData);
            } catch (IllegalAccessException | InvocationTargetException e) {
                e.printStackTrace();
            }
        }

    }

    private void writeToCell(Cell cell, Object data) {
        if (data instanceof String) {
            cell.setCellValue((String) data);
            cell.setCellType(CellType.STRING);
            return;
        }
        if (data instanceof Integer) {
            int num = (int) data;
            cell.setCellValue(num);
            cell.setCellType(CellType.NUMERIC);
            return;
        }
        if (data instanceof Float) {
            float num = (float) data;
            cell.setCellValue(num);
            cell.setCellType(CellType.NUMERIC);
        }
        if (data instanceof Double) {
            double num = (double) data;
            cell.setCellValue(num);
            cell.setCellType(CellType.NUMERIC);
        }

    }


    @Getter
    @Setter
    @NoArgsConstructor
    @AllArgsConstructor
    static
    class Column {
        private int number;
        private String header;
        private String name;
        private Method methodToSet;
    }

}
