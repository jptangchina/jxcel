package com.jptangchina.jxcel;

import com.google.common.base.Preconditions;
import com.google.common.base.Strings;
import com.jptangchina.jxcel.annotation.JxcelCell;
import com.jptangchina.jxcel.annotation.JxcelSheet;
import com.jptangchina.jxcel.exception.JxcelGenerateException;
import lombok.AllArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;

import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;

@Slf4j
@AllArgsConstructor
public class JxcelGenrator {

    private Workbook workbook;

    public Workbook generateWorkbook(List<?> data) {
        Preconditions.checkArgument(CollectionUtils.isNotEmpty(data), "Data must not be empty.");
        Class<?> type = data.get(0).getClass();
        Map<String, Field> headers = getSheetHeaderWithOrder(type);
        Preconditions.checkArgument(headers.size() > 0, "No available field.");
        CellStyle headerStyle = createHeaderStyle(type, workbook);
        Sheet sheet = createSheet(type, workbook);
        AtomicInteger cellIndex = new AtomicInteger(0);
        AtomicInteger rowIndex = new AtomicInteger(1);
        createHeaderRow(sheet, headers, cellIndex, headerStyle);
        createDataRows(data, headers, cellIndex, rowIndex, sheet);
        resizeColumn(headers, sheet);
        return workbook;
    }

    public void generateFile(List<?> data, String outputPath) {
        try(FileOutputStream fos = new FileOutputStream(outputPath);
            Workbook workbook = generateWorkbook(data)) {
            workbook.write(fos);
            fos.flush();
        } catch (IOException e) {
            log.error(e.getMessage(), e);
            throw new JxcelGenerateException("File generate failed.");
        }
    }

    private Map<String, Field> getSheetHeaderWithOrder(Class<?> type) {
        Map<String, Field> result = new LinkedHashMap<>();
        List<Field> fields = Arrays.stream(type.getDeclaredFields())
            .filter(f -> f.isAnnotationPresent(JxcelCell.class))
            .sorted(Comparator.comparingInt(f -> f.getAnnotation(JxcelCell.class).order()))
            .collect(Collectors.toList());
        fields.forEach(f -> result.put(f.getName(), f));
        return result;
    }

    private void setStringCellValue(Object obj, Cell cell, Field field) throws IllegalAccessException {
        JxcelCell jxcelCell = field.getAnnotation(JxcelCell.class);
        field.setAccessible(true);
        Object value = field.get(obj);
        // 日期格式处理
        if(field.getType() == Date.class) {
            if (value == null) {
                value = "";
            } else {
                DateTime time = new DateTime(value);
                if(!Strings.isNullOrEmpty(jxcelCell.format())) {
                    value = time.toString(jxcelCell.format());
                }
            }
        }
        // 内容转换
        if(jxcelCell.parse().length > 0) {
            if (Boolean.class.equals(field.getType()) || Boolean.TYPE.equals(field.getType())) {
                int index = Boolean.parseBoolean(value.toString()) ? 1 : 0;
                value = jxcelCell.parse()[index];
            }else {
                int index = (int)value;
                value = jxcelCell.parse()[index];
            }
        }
        // 后缀拼接
        value += jxcelCell.suffix();
        cell.setCellValue(value.toString());
    }

    private Sheet createSheet(Class<?> type, Workbook workbook) {
        Sheet sheet;
        JxcelSheet jxcelSheet = type.getAnnotation(JxcelSheet.class);
        if(null != jxcelSheet) {
            sheet = Strings.isNullOrEmpty(jxcelSheet.value()) ?
                workbook.createSheet() : workbook.createSheet(jxcelSheet.value());
        }else {
            sheet = workbook.createSheet();
        }
        return sheet;
    }

    private CellStyle createHeaderStyle(Class<?> type, Workbook workbook) {
        JxcelSheet jxcelSheet = type.getAnnotation(JxcelSheet.class);
        CellStyle style = workbook.createCellStyle();
        if(null == jxcelSheet) {
            return style;
        }
        if(jxcelSheet.color() > -1) {
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            style.setFillForegroundColor(jxcelSheet.color());
        }
        return style;
    }

    private void createDataRows(List<?> data, Map<String, Field> headers, AtomicInteger cellIndex,
        AtomicInteger rowIndex, Sheet sheet) {
        data.forEach(d -> {
            Row row = sheet.createRow(rowIndex.getAndIncrement());
            headers.values().forEach(f -> {
                Cell cell = row.createCell(cellIndex.getAndIncrement(), CellType.STRING);
                try {
                    setStringCellValue(d, cell, f);
                } catch (IllegalAccessException e) {
                    log.error(e.getMessage(), e);
                    throw new JxcelGenerateException("Can not get value.");
                }
            });
            cellIndex.set(0);
        });
    }

    private void createHeaderRow(Sheet sheet, Map<String, Field> headers, AtomicInteger cellIndex, CellStyle style) {
        Row head = sheet.createRow(0);
        headers.forEach((k, v) -> {
            Cell cell = head.createCell(cellIndex.getAndIncrement(), CellType.STRING);
            JxcelCell jxcelCell = v.getAnnotation(JxcelCell.class);
            String headName = Strings.isNullOrEmpty(jxcelCell.value()) ? v.getName() : jxcelCell.value();
            cell.setCellStyle(style);
            cell.setCellValue(headName);
        });
        cellIndex.set(0);
    }

    private void resizeColumn(Map<String, Field> headers, Sheet sheet) {
        int index = 0;
        for(Field field : headers.values()) {
            JxcelCell jxcelCell = field.getAnnotation(JxcelCell.class);
            if(jxcelCell.autoResize()) {
                sheet.autoSizeColumn(index, true);
            }
            index ++;
        }
    }

    public static JxcelGenrator xlsGenrator() {
        return new JxcelGenrator(new HSSFWorkbook());
    }

    public static JxcelGenrator xlsxGenrator() {
        return new JxcelGenrator(new XSSFWorkbook());
    }

}
