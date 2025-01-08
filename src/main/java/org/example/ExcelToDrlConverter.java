package org.example;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.stream.Stream;

public class ExcelToDrlConverter {

    public static void main(String[] args) {
        if (args.length < 1) {
            System.out.println("Usage: java ExcelToDrlConverter <excel-file-path> ");
            return;
        }

        String excelFilePath = args[0];

        try {
            convertExcelToDrl(excelFilePath);
            System.out.println("DRL files generated successfully.");
        } catch (IOException e) {
            System.err.println("Error during conversion: " + e.getMessage());
            e.printStackTrace();
        }
    }

    public static void convertExcelToDrl(String excelFilePath) throws IOException {
        Path excelDirectory = Paths.get(excelFilePath);

        try (Stream<Path> paths = Files.walk(excelDirectory)) {
            paths.filter(Files::isRegularFile)
                    .filter(path -> path.toString().endsWith(".xls"))
                    .forEach(path -> {
                        try {
                            System.out.println("Processing file: {}" + path);

                            byte[] excelData = Files.readAllBytes(path);

                            // Convert to DRL

                            String drlContent;
                            if (StringUtils.equalsIgnoreCase("DraftWorkflowRules.xls", path.getFileName().toString())) {
                                drlContent = excelToDrlConverter(excelData, "DraftWorkflowParams");
                            }else {
                                drlContent = excelToDrlConverter(excelData, "WorkflowParams");
                            }

                            System.out.println("Generated DRL:\n" + drlContent);

                        } catch (Exception e) {
                            System.err.println("Failed to process file: " + path + " - " + e.getMessage());
                        }
                    });
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static String excelToDrlConverter(byte[] excelFile, String drlClass) throws IOException {
        StringBuilder drlContent = new StringBuilder();

        drlContent.append("package com.order.rules;\n\n")
                .append("import com.order.rules.").append(drlClass).append(";\n\n");

        try (InputStream fis = new ByteArrayInputStream(excelFile);
             Workbook workbook = new HSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);

            // Locate header row
            int headerRowIndex = StringUtils.equalsIgnoreCase(drlClass, "DraftWorkflowParams") ?
                    findHeaderRow(sheet, "Circle", "Rule Key", "Workflow Name", "Workflow Version") :
                    findHeaderRow(sheet, "Circle", "Rule Key", "Workflow Name", "Workflow Version", "Channel");
            if (headerRowIndex == -1) {
                throw new IllegalArgumentException("Header row not found");
            }

            // Get column indexes for the required fields
            Row headerRow = sheet.getRow(headerRowIndex);
            int circleIndex = findColumnIndex(headerRow, "Circle");
            int ruleKeyIndex = findColumnIndex(headerRow, "Rule Key");
            int workflowNameIndex = findColumnIndex(headerRow, "Workflow Name");
            int workflowVersionIndex = findColumnIndex(headerRow, "Workflow Version");
            int channelIndex = 0;
            if (StringUtils.equalsIgnoreCase(drlClass, "WorkflowParams")) {
                channelIndex = findColumnIndex(headerRow, "Channel");
            }

            // Process rows below the header
            for (int i = headerRowIndex + 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null || isRowEmpty(row)) {
                    continue;
                }

                String circle = getCellValue(row.getCell(circleIndex));
                String ruleKey = getCellValue(row.getCell(ruleKeyIndex));
                String workflowName = getCellValue(row.getCell(workflowNameIndex));
                String workflowVersion = getCellValue(row.getCell(workflowVersionIndex));
                String channel = StringUtils.equalsIgnoreCase(drlClass, "WorkflowParams") ?
                        getCellValue(row.getCell(channelIndex)) : null;

                // Generate DRL string
                String drl = StringUtils.equalsIgnoreCase(drlClass, "DraftWorkflowParams") ?
                        generateDrlStringForDraftWorkflow(i, circle, ruleKey, workflowName, workflowVersion) :
                        generateDrlStringForOrderWorkflow(i, circle, ruleKey, workflowName, workflowVersion, channel);
                drlContent.append(drl);
            }
        } catch (IOException e) {
            throw new RuntimeException("Failed to convert Excel to DRL: " + e.getMessage(), e);
        }

        return drlContent.toString();
    }

    private static int findHeaderRow(Sheet sheet, String... headers) {
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row != null && containsHeaders(row, headers)) {
                return i;
            }
        }
        return -1;
    }

    private static boolean containsHeaders(Row row, String... headers) {
        for (String header : headers) {
            if (findColumnIndex(row, header) == -1) {
                return false;
            }
        }
        return true;
    }

    private static int findColumnIndex(Row row, String header) {
        for (int i = 0; i < row.getLastCellNum(); i++) {
            if (header.equalsIgnoreCase(getCellValue(row.getCell(i)))) {
                return i;
            }
        }
        return -1;
    }

    private static boolean isRowEmpty(Row row) {
        for (int i = 0; i < row.getLastCellNum(); i++) {
            Cell cell = row.getCell(i);
            if (cell != null && cell.getCellType() != CellType.BLANK) {
                return false;
            }
        }
        return true;
    }

    private static String getCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }
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

    private static String generateDrlStringForDraftWorkflow(int ruleIndex,
                                                            String circle,
                                                            String ruleKey,
                                                            String workflowName,
                                                            String workflowVersion) {
        return String.format(
                "// rule values at row %d\n" +
                        "rule \"Draft Rule Config_%d\"\n" +
                        "    when\n" +
                        "        fact : com.order.rules.DraftWorkflowParams(circle in (%s), ruleKey == \"%s\")\n" +
                        "    then\n" +
                        "        fact.setWorkflowName(\"%s\");\n" +
                        "        fact.setWorkflowVersion(\"%s\");\n" +
                        "end\n\n",
                ruleIndex + 1, ruleIndex, circle, ruleKey, workflowName, workflowVersion
        );
    }

    private static String generateDrlStringForOrderWorkflow(int ruleIndex,
                                                            String circle,
                                                            String ruleKey,
                                                            String workflowName,
                                                            String workflowVersion,
                                                            String channel) {
        return String.format(
                "// rule values at row %d\n" +
                        "rule \"Rule Config_%d\"\n" +
                        "    when\n" +
                        "        fact : com.order.rules.WorkflowParams(channel in (%s), circle in (%s), ruleKey == \"%s\")\n" +
                        "    then\n" +
                        "        fact.setWorkflowName(\"%s\");\n" +
                        "        fact.setWorkflowVersion(\"%s\");\n" +
                        "end\n\n",
                ruleIndex + 1, ruleIndex, channel, circle, ruleKey, workflowName, workflowVersion
        );
    }

}
