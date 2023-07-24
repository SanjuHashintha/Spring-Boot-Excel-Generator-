package com.example.demo2;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartUtils;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.plot.PiePlot;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.chart.plot.RingPlot;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.data.general.DefaultPieDataset;

import java.io.FileInputStream;
import java.io.FileOutputStream;

public class ExcelGenerator{

    public static void generateExcelFile() throws Exception {
        // Read the input Excel file
        FileInputStream fileIn = new FileInputStream("input.xlsx");
        Workbook inputWorkbook = new XSSFWorkbook(fileIn);
        Sheet inputSheet = inputWorkbook.getSheetAt(0);

        // Create a new workbook for the output
        Workbook outputWorkbook = new XSSFWorkbook();
        Sheet outputSheet = outputWorkbook.createSheet("Data");

        // Get the table head from the input sheet and copy to the output sheet
        Row headRow = inputSheet.getRow(0);
        Row outputHeadRow = outputSheet.createRow(0);
        for (Cell headCell : headRow) {
            Cell outputHeadCell = outputHeadRow.createCell(headCell.getColumnIndex());
            outputHeadCell.setCellValue(headCell.getStringCellValue());

            // Apply cell style for table head
            CellStyle headCellStyle = outputWorkbook.createCellStyle();
            headCellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            headCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headCellStyle.setBorderBottom(BorderStyle.THIN);
            headCellStyle.setBorderTop(BorderStyle.THIN);
            headCellStyle.setBorderLeft(BorderStyle.THIN);
            headCellStyle.setBorderRight(BorderStyle.THIN);
            headCellStyle.setAlignment(HorizontalAlignment.CENTER);
            Font headFont = outputWorkbook.createFont();
            headFont.setBold(true);
            headCellStyle.setFont(headFont);
            outputHeadCell.setCellStyle(headCellStyle);
        }

        // Copy the data from the input sheet to the output sheet
        int rowCount = inputSheet.getLastRowNum();
        for (int i = 1; i <= rowCount; i++) {
            Row inputRow = inputSheet.getRow(i);
            Row outputRow = outputSheet.createRow(i);

            for (Cell inputCell : inputRow) {
                int a[] = {10,5,2,8,3,4};
                Cell outputCell = outputRow.createCell(inputCell.getColumnIndex());
                CellType inputCellType = inputCell.getCellType();

                if (inputCellType == CellType.STRING) {
                    outputCell.setCellValue(inputCell.getStringCellValue());
                } else if (inputCellType == CellType.NUMERIC) {
                    outputCell.setCellValue(inputCell.getNumericCellValue());
                } else if (inputCellType == CellType.BOOLEAN) {
                    outputCell.setCellValue(inputCell.getBooleanCellValue());
                } else if (inputCellType == CellType.BLANK) {
                    outputCell.setCellValue(a[i-1]);

                }

                // Apply cell style for data cells
                CellStyle dataCellStyle = outputWorkbook.createCellStyle();
                dataCellStyle.setBorderBottom(BorderStyle.THIN);
                dataCellStyle.setBorderTop(BorderStyle.THIN);
                dataCellStyle.setBorderLeft(BorderStyle.THIN);
                dataCellStyle.setBorderRight(BorderStyle.THIN);
                outputCell.setCellStyle(dataCellStyle);
            }
        }

        // Generate the pie chart
        DefaultPieDataset pieDataset = new DefaultPieDataset();
        DefaultCategoryDataset barDataset = new DefaultCategoryDataset();
        for (int i = 1; i <= rowCount; i++) {
            Row dataRow = outputSheet.getRow(i);
            Cell genderCell = dataRow.getCell(0);
            Cell countCell = dataRow.getCell(1);
            String gender = genderCell != null ? genderCell.getStringCellValue() : "";
            double count = countCell != null ? countCell.getNumericCellValue() : 0.0;
            pieDataset.setValue(gender, count);
            barDataset.addValue(count, "Count", gender);
        }

        JFreeChart pieChart = ChartFactory.createPieChart("Gender Distribution (Pie Chart)", pieDataset, true, true, false);
        JFreeChart barChart = ChartFactory.createBarChart("Gender Distribution (Bar Chart)", "Gender", "Count", barDataset, PlotOrientation.VERTICAL, true, true, false);
        JFreeChart donutChart = createDonutChart(pieDataset);

        // Convert the charts to images
        byte[] pieChartImageBytes = ChartUtils.encodeAsPNG(pieChart.createBufferedImage(400, 300));
        byte[] barChartImageBytes = ChartUtils.encodeAsPNG(barChart.createBufferedImage(400, 300));
        byte[] donutChartImageBytes = ChartUtils.encodeAsPNG(donutChart.createBufferedImage(400, 300));

        // Create sheets for the charts and add the images
        Sheet pieChartSheet = outputWorkbook.createSheet("Pie Chart");
        Sheet barChartSheet = outputWorkbook.createSheet("Bar Chart");
        Sheet donutChartSheet = outputWorkbook.createSheet("Donut Chart");
        Drawing<?> pieChartDrawing = pieChartSheet.createDrawingPatriarch();
        Drawing<?> barChartDrawing = barChartSheet.createDrawingPatriarch();
        Drawing<?> donutChartDrawing = donutChartSheet.createDrawingPatriarch();
        ClientAnchor pieChartAnchor = pieChartDrawing.createAnchor(0, 0, 0, 0, 0, 2, 7, 20);
        ClientAnchor barChartAnchor = barChartDrawing.createAnchor(0, 0, 0, 0, 0, 2, 7, 20);
        ClientAnchor donutChartAnchor = donutChartDrawing.createAnchor(0, 0, 0, 0, 0, 2, 7, 20);
        int pieChartPictureIdx = outputWorkbook.addPicture(pieChartImageBytes, Workbook.PICTURE_TYPE_PNG);
        int barChartPictureIdx = outputWorkbook.addPicture(barChartImageBytes, Workbook.PICTURE_TYPE_PNG);
        int donutChartPictureIdx = outputWorkbook.addPicture(donutChartImageBytes, Workbook.PICTURE_TYPE_PNG);
        Picture pieChartPicture = pieChartDrawing.createPicture(pieChartAnchor, pieChartPictureIdx);
        Picture barChartPicture = barChartDrawing.createPicture(barChartAnchor, barChartPictureIdx);
        Picture donutChartPicture = donutChartDrawing.createPicture(donutChartAnchor, donutChartPictureIdx);
        pieChartPicture.resize();
        barChartPicture.resize();
        donutChartPicture.resize();

        // Apply table borders to the output sheet
        for (Row row : outputSheet) {
            for (Cell cell : row) {
                CellStyle cellStyle = outputWorkbook.createCellStyle();
                cellStyle.setBorderBottom(BorderStyle.THIN);
                cellStyle.setBorderTop(BorderStyle.THIN);
                cellStyle.setBorderLeft(BorderStyle.THIN);
                cellStyle.setBorderRight(BorderStyle.THIN);
                cell.setCellStyle(cellStyle);
            }
        }

        // Write the output workbook to a file
        FileOutputStream fileOut = new FileOutputStream("output.xlsx");
        outputWorkbook.write(fileOut);
        fileOut.close();

        // Close the workbooks
        outputWorkbook.close();
        inputWorkbook.close();
    }

    private static JFreeChart createDonutChart(DefaultPieDataset dataset) {
        JFreeChart chart = ChartFactory.createRingChart("Gender Distribution (Donut Chart)", dataset, true, true, false);
        RingPlot plot = (RingPlot) chart.getPlot();
        plot.setLabelGenerator(null);
        plot.setSectionDepth(0.3);
        return chart;
    }

}
