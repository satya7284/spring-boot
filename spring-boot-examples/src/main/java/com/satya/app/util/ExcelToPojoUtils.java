package com.satya.app.util;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Logger;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelToPojoUtils {
	
	private final static Logger LOGGER = Logger.getLogger(ExcelToPojoUtils.class.getName());
	
    public static final String BOOLEAN_TRUE = "1";
    public static final String MERGED_HEADER_NAME = "header1";

    private static String strToFieldName(String str) {
        str = str.replaceAll("[^a-zA-Z0-9]", "");
        return str.length() > 0 ? str.substring(0, 1).toLowerCase() + str.substring(1) : null;
    }

    public static <T> List<T> toPojoList(Class<T> type, InputStream inputStream) throws IOException {
        List<T> results = new ArrayList<>();
        Workbook workbook = WorkbookFactory.create(inputStream);
        Sheet sheet = workbook.getSheetAt(0);

        // header column names
        List<String> colNames = new ArrayList<>();
        Row headerRow = sheet.getRow(0);
        for (int i = 0; i < headerRow.getPhysicalNumberOfCells(); i++) {
            Cell headerCell = headerRow.getCell(i, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
            colNames.add(headerCell != null ? strToFieldName(headerCell.getStringCellValue()) : null);
        }

        String mergedValue = null;
        for (int j = 1; j < sheet.getPhysicalNumberOfRows(); j++) {
            Row row = sheet.getRow(j);
            try {
                T result = type.getDeclaredConstructor().newInstance();
                for (int k = 0; k < row.getPhysicalNumberOfCells(); k++) {
                	String headerName = colNames.get(k);
                    if (headerName != null) {
                        Cell cell = row.getCell(k, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                        if (cell != null) {
                            DataFormatter formatter = new DataFormatter();
                            String strValue = formatter.formatCellValue(cell);
                            
                            if(MERGED_HEADER_NAME.equals(headerName)) {
                            	if(strValue != null) {
                            		mergedValue = strValue;
                            	} else {
                            		strValue = mergedValue;
                            	}
                            }
                            
                            Field field = type.getDeclaredField(colNames.get(k));
                            field.setAccessible(true);
                            if (field != null) {
                                Object value = null;
                                if (field.getType().equals(Long.class)) {
                                    value = Long.valueOf(strValue);
                                } else if (field.getType().equals(String.class)) {
                                    value = cell.getStringCellValue();
                                } else if (field.getType().equals(Integer.class)) {
                                    value = Integer.valueOf(strValue);
                                } else if (field.getType().equals(LocalDate.class)) {
                                    value = LocalDate.parse(strValue);
                                } else if (field.getType().equals(LocalDateTime.class)) {
                                    value = LocalDateTime.parse(strValue);
                                } else if (field.getType().equals(Boolean.class)) {
                                    value = BOOLEAN_TRUE.equals(strValue);
                                } else if (field.getType().equals(BigDecimal.class)) {
                                    value = new BigDecimal(strValue);
                                }
                                field.set(result, value);
                            }
                        }
                    }
                }
                boolean allNulls = NullChecker.allNull(result);
                if(!allNulls) {
                	Field field = type.getDeclaredField(MERGED_HEADER_NAME);
                    field.setAccessible(true);
                	if(field.get(result) == null && StringUtils.isNotEmpty(mergedValue)) {
                		field.set(result, mergedValue);
                	}
                	results.add(result);
                }
            } catch (Exception e) {
                LOGGER.info(e.getMessage());
            }
        }

        return results;
    }
}