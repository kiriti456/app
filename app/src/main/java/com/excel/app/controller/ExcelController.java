package com.excel.app.controller;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

@CrossOrigin("https://main--excel-uofm.netlify.app/excel-operations")
@RestController
@RequestMapping("/api/excel")
public class ExcelController {
	
	@GetMapping("hello-world")
	public String hello() {
		return "Hello World!!";
	}

    @PostMapping("/merge")
    public ResponseEntity<byte[]> mergeExcelFiles(
            @RequestBody MultipartFile[] files,
            @RequestParam String columnName) throws IOException {

    	if (files.length != 2) {
            return ResponseEntity.badRequest().body(null);
        }

        Workbook workbook1 = new XSSFWorkbook(files[0].getInputStream());
        Workbook workbook2 = new XSSFWorkbook(files[1].getInputStream());

        Sheet sheet1 = workbook1.getSheetAt(0);
        Sheet sheet2 = workbook2.getSheetAt(0);

        int columnIndex1 = getColumnIndex(sheet1, columnName);
        int columnIndex2 = getColumnIndex(sheet2, columnName);

        if (columnIndex1 == -1 || columnIndex2 == -1) {
            return ResponseEntity.status(HttpStatus.BAD_REQUEST).body(null);
        }

        Map<String, Row> rowsSheet1 = new HashMap<>();
        Map<String, Row> rowsSheet2 = new HashMap<>();

        // Extract headers and rows
        List<String> headers1 = getHeaders(sheet1);
        List<String> headers2 = getHeaders(sheet2);

        Set<String> allHeaders = new LinkedHashSet<>(headers1);
        allHeaders.addAll(headers2);

        List<String> allHeadersList = new ArrayList<>(allHeaders);

        // Populate row maps
        populateRowMap(sheet1, columnIndex1, rowsSheet1);
        populateRowMap(sheet2, columnIndex2, rowsSheet2);

        // Create workbooks for union and intersection
        Workbook unionWorkbook = new XSSFWorkbook();
        Workbook intersectionWorkbook = new XSSFWorkbook();
        Sheet unionSheet = unionWorkbook.createSheet("Union");
        Sheet intersectionSheet = intersectionWorkbook.createSheet("Intersection");

        // Write headers
        writeHeaders(allHeadersList, unionSheet);
        writeHeaders(allHeadersList, intersectionSheet);

        // Write union data
        Set<String> allKeys = new LinkedHashSet<>(rowsSheet1.keySet());
        allKeys.addAll(rowsSheet2.keySet());

        Map<String, Row> unionMap = new LinkedHashMap<>();
        for (String key : allKeys) {
            Row row1 = rowsSheet1.get(key);
            Row row2 = rowsSheet2.get(key);
            Row unionRow = unionSheet.createRow(unionSheet.getLastRowNum() + 1);
            if (row1 != null) {
                copyRowData(row1, unionRow, allHeadersList);
            }
            if (row2 != null) {
                copyRowData(row2, unionRow, allHeadersList);
            }
            unionMap.put(key, unionRow);
        }

        // Write intersection data
        Set<String> intersectionKeys = new HashSet<>(rowsSheet1.keySet());
        intersectionKeys.retainAll(rowsSheet2.keySet());

        for (String key : intersectionKeys) {
            Row row1 = rowsSheet1.get(key);
            Row row2 = rowsSheet2.get(key);
            Row intersectionRow = intersectionSheet.createRow(intersectionSheet.getLastRowNum() + 1);
            if (row1 != null) {
                copyRowData(row1, intersectionRow, allHeadersList);
            }
            if (row2 != null) {
                copyRowData(row2, intersectionRow, allHeadersList);
            }
        }

        // Convert workbooks to byte arrays
        byte[] unionExcelData = convertWorkbookToByteArray(unionWorkbook);
        byte[] intersectionExcelData = convertWorkbookToByteArray(intersectionWorkbook);

        // Prepare response as a zip file containing both Excel files
        byte[] zipData = createZipFile(unionExcelData, intersectionExcelData);

        HttpHeaders headers = new HttpHeaders();
        headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=merged_intersected_files.zip");
        headers.add(HttpHeaders.CONTENT_TYPE, "application/zip");

        return new ResponseEntity<>(zipData, headers, HttpStatus.OK);
    }

    private int getColumnIndex(Sheet sheet, String columnName) {
        Row headerRow = sheet.getRow(0);
        for (Cell cell : headerRow) {
            if (cell.getStringCellValue().equalsIgnoreCase(columnName)) {
                return cell.getColumnIndex();
            }
        }
        return -1;
    }

    private List<String> getHeaders(Sheet sheet) {
        Row headerRow = sheet.getRow(0);
        List<String> headers = new ArrayList<>();
        for (Cell cell : headerRow) {
            headers.add(cell.getStringCellValue());
        }
        return headers;
    }

    private void populateRowMap(Sheet sheet, int columnIndex, Map<String, Row> rowMap) {
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                Cell cell = row.getCell(columnIndex);
                String key = normalizeKey(getCellValueAsString(cell));
                rowMap.put(key, row);
            }
        }
    }

    private String getCellValueAsString(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            default:
                return "";
        }
    }

    private String normalizeKey(String key) {
        if (key == null) {
            return "";
        }
        return key.replaceAll("[^a-zA-Z0-9]", "").toLowerCase().trim();
    }

    private void writeHeaders(List<String> headers, Sheet sheet) {
        Row row = sheet.createRow(0);
        for (int i = 0; i < headers.size(); i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(headers.get(i));
        }
    }

    private void copyRowData(Row sourceRow, Row targetRow, List<String> allHeaders) {
        Map<Integer, Cell> cellMap = new HashMap<>();
        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            Cell cell = sourceRow.getCell(i);
            if (cell != null) {
                cellMap.put(i, cell);
            }
        }
        for (int i = 0; i < allHeaders.size(); i++) {
            Cell targetCell = targetRow.createCell(i);
            String header = allHeaders.get(i);
            Integer columnIndex = getColumnIndexByHeader(sourceRow.getSheet(), header);
            if (columnIndex != null && cellMap.containsKey(columnIndex)) {
                Cell sourceCell = cellMap.get(columnIndex);
                switch (sourceCell.getCellType()) {
                    case STRING:
                        targetCell.setCellValue(sourceCell.getStringCellValue());
                        break;
                    case NUMERIC:
                        targetCell.setCellValue(sourceCell.getNumericCellValue());
                        break;
                    case BOOLEAN:
                        targetCell.setCellValue(sourceCell.getBooleanCellValue());
                        break;
                    case FORMULA:
                        targetCell.setCellFormula(sourceCell.getCellFormula());
                        break;
                    default:
                        targetCell.setCellValue("");
                        break;
                }
            } else {
                targetCell.setCellValue("");
            }
        }
    }

    private Integer getColumnIndexByHeader(Sheet sheet, String header) {
        Row headerRow = sheet.getRow(0);
        for (Cell cell : headerRow) {
            if (cell.getStringCellValue().equalsIgnoreCase(header)) {
                return cell.getColumnIndex();
            }
        }
        return null;
    }

    private byte[] convertWorkbookToByteArray(Workbook workbook) throws IOException {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        workbook.write(outputStream);
        return outputStream.toByteArray();
    }

    private byte[] createZipFile(byte[] unionExcelData, byte[] intersectionExcelData) throws IOException {
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        try (java.util.zip.ZipOutputStream zipOut = new java.util.zip.ZipOutputStream(byteArrayOutputStream)) {
            // Add union file to the zip
            java.util.zip.ZipEntry zipEntry1 = new java.util.zip.ZipEntry("union.xlsx");
            zipOut.putNextEntry(zipEntry1);
            zipOut.write(unionExcelData);
            zipOut.closeEntry();

            // Add intersection file to the zip
            java.util.zip.ZipEntry zipEntry2 = new java.util.zip.ZipEntry("intersection.xlsx");
            zipOut.putNextEntry(zipEntry2);
            zipOut.write(intersectionExcelData);
            zipOut.closeEntry();
        }

        return byteArrayOutputStream.toByteArray();
    }
    
    
    //-----------------------------------------------------------------------
    
    
    
    @PostMapping("/split")
    public ResponseEntity<byte[]> splitExcelFileByColumn(
            @RequestBody MultipartFile file,
            @RequestParam String columnName) throws IOException {

        Workbook workbook = new XSSFWorkbook(file.getInputStream());
        Sheet sheet = workbook.getSheetAt(0);

        int columnIndex = getColumnIndex(sheet, columnName);

        if (columnIndex == -1) {
            return ResponseEntity.status(HttpStatus.BAD_REQUEST).body(null);
        }

        // Map to hold rows grouped by the unique values in the specified column
        Map<String, List<Row>> groupedRows = new LinkedHashMap<>();

        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                Cell cell = row.getCell(columnIndex);
                String key = normalizeKey(getCellValueAsString(cell));
                groupedRows.computeIfAbsent(key, k -> new ArrayList<>()).add(row);
            }
        }

        // Prepare zip file with individual files
        byte[] zipData = createZipFileFromGroupedRows(groupedRows, sheet.getRow(0));

        HttpHeaders headers = new HttpHeaders();
        headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=split_files.zip");
        headers.add(HttpHeaders.CONTENT_TYPE, "application/zip");

        return new ResponseEntity<>(zipData, headers, HttpStatus.OK);
    }

    private byte[] createZipFileFromGroupedRows(Map<String, List<Row>> groupedRows, Row headerRow) throws IOException {
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        try (java.util.zip.ZipOutputStream zipOut = new java.util.zip.ZipOutputStream(byteArrayOutputStream)) {
            for (Map.Entry<String, List<Row>> entry : groupedRows.entrySet()) {
                String key = entry.getKey();
                List<Row> rows = entry.getValue();

                Workbook newWorkbook = new XSSFWorkbook();
                Sheet newSheet = newWorkbook.createSheet("Sheet1");

                // Copy header row
                copyRowData(headerRow, newSheet.createRow(0), headerRow);

                // Copy rows to new sheet
                for (int i = 0; i < rows.size(); i++) {
                    Row newRow = newSheet.createRow(i + 1);
                    copyRowData(rows.get(i), newRow, headerRow);
                }

                // Convert workbook to byte array
                ByteArrayOutputStream tempOutputStream = new ByteArrayOutputStream();
                newWorkbook.write(tempOutputStream);
                byte[] excelData = tempOutputStream.toByteArray();

                // Add to zip
                java.util.zip.ZipEntry zipEntry = new java.util.zip.ZipEntry(key + ".xlsx");
                zipOut.putNextEntry(zipEntry);
                zipOut.write(excelData);
                zipOut.closeEntry();
            }
        }
        return byteArrayOutputStream.toByteArray();
    }

    private void copyRowData(Row sourceRow, Row targetRow, Row headerRow) {
        for (int i = 0; i < headerRow.getLastCellNum(); i++) {
            Cell targetCell = targetRow.createCell(i);
            Cell sourceCell = sourceRow.getCell(i);
            if (sourceCell != null) {
                switch (sourceCell.getCellType()) {
                    case STRING:
                        targetCell.setCellValue(sourceCell.getStringCellValue());
                        break;
                    case NUMERIC:
                        targetCell.setCellValue(sourceCell.getNumericCellValue());
                        break;
                    case BOOLEAN:
                        targetCell.setCellValue(sourceCell.getBooleanCellValue());
                        break;
                    case FORMULA:
                        targetCell.setCellFormula(sourceCell.getCellFormula());
                        break;
                    default:
                        targetCell.setCellValue("");
                        break;
                }
            } else {
                targetCell.setCellValue("");
            }
        }
    }
    
}