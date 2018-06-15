package sample;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javafx.application.Application;
import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

public class Controller extends Application {

    Stage stage;

    /**
     * TRKA - CONTENDERS
     */
    Map<String, List<String>> races = new HashMap<>();

    Map<String, List<String>> treeMap;

    String[] time = {"11:00"};


    @FXML
    private void browseButton(ActionEvent event) {

        System.out.println("wtf?");


        FileChooser fileChooser = new FileChooser();

        fileChooser.getExtensionFilters().addAll(
                new FileChooser.ExtensionFilter("Excel", "*.xlsx"));
        fileChooser.setTitle("Open Resource File");
        List<File> entries = fileChooser.showOpenMultipleDialog(stage);
        String club = null;
        /**
         * CONTENDERS - CLUB
         */
        List<String> contendersClub = new ArrayList<>();
        String trka = null;
        List<String> contenders = new ArrayList<>();

        for (File f : entries) {
            try {

                club = f.getName();

                XSSFWorkbook wb = new XSSFWorkbook(f);
                //POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(f));
                //HSSFWorkbook wb = new HSSFWorkbook(fs);
                XSSFSheet sheet = wb.getSheetAt(0);
                XSSFRow row;
                XSSFCell cell;

                int rows; // No of rows
                rows = sheet.getPhysicalNumberOfRows();

                int cols = 0; // No of columns
                int tmp = 0;

                // This trick ensures that we get the data properly even if it doesn't start from first few rows
                for (int i = 0; i < 10 || i < rows; i++) {
                    row = sheet.getRow(i);
                    if (row != null) {
                        tmp = sheet.getRow(i).getPhysicalNumberOfCells();
                        if (tmp > cols) cols = tmp;
                    }
                }

                for (int r = 6; r < 30; r++) {
                    row = sheet.getRow(r);
                    if (row != null) {
                        cell = row.getCell(2);
                        if (cell != null) {
                            // Your code here
                            System.out.println(cell);

                            Integer no = r - 5;

                            if (no < 10) {
                                trka = "0" + no.toString() + ">";
                            } else {
                                trka = no.toString() + ">";
                            }


                            trka += cell.toString();

                            if (row.getCell(3) != null) {
                                trka += ">" + row.getCell(3);

                            }


                            if (!"".equals(row.getCell(4).toString())) {
                                String raw = row.getCell(4).toString();

                                contenders = Arrays.asList(raw.split(","));

                                for (String s :
                                        contenders) {

                                    if (races.get(trka) != null) {
                                        contendersClub = races.get(trka);
                                    }
                                    contendersClub.add(s + ">" + club.replace(".xlsx", ""));

                                }


                            } else {
                                continue;
                            }


                        }
                    }
                    races.put(trka, contendersClub);
                    contendersClub = new ArrayList<String>();
                    contenders = new ArrayList<>();
                }

            wb.close();
            } catch (Exception ioe) {
                ioe.printStackTrace();
            }
        }

        treeMap = new TreeMap<String, List<String>>(races);
        makeOutput();
    }

//
//    private static final String[] titles = {
//            "Person",	"ID", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun",
//            "Total\nHrs", "Overtime\nHrs", "Regular\nHrs"
//    };
//
//    private static Object[][] sample_data = {
//            {"Yegor Kozlov", "YK", 5.0, 8.0, 10.0, 5.0, 5.0, 7.0, 6.0},
//            {"Gisella Bronzetti", "GB", 4.0, 3.0, 1.0, 3.5, null, null, 4.0},
//    };

    void makeOutput() {
        Workbook workbook = new XSSFWorkbook(); // new HSSFWorkbook() for generating `.xls` file

        /* CreationHelper helps us create instances of various things like DataFormat,
           Hyperlink, RichTextString etc, in a format (HSSF, XSSF) independent way */
        CreationHelper createHelper = workbook.getCreationHelper();

        // Create a Sheet
        Sheet sheet = workbook.createSheet("Employee");

        // Create a Font for styling header cells
        //Font headerFont = workbook.createFont();
        //headerFont.setBold(true);
        //headerFont.setFontHeightInPoints((short) 14);
        //headerFont.setColor(IndexedColors.RED.getIndex());

        // Create a CellStyle with the font
        //CellStyle headerCellStyle = workbook.createCellStyle();
        //headerCellStyle.setFont(headerFont);


        Font headLine = workbook.createFont();
        headLine.setBold(true);
        headLine.setFontHeightInPoints((short) 12);
        //headerFont.setColor(IndexedColors.RED.getIndex());

        CellStyle headLineCellStyle = workbook.createCellStyle();
        headLineCellStyle.setFont(headLine);
        headLineCellStyle.setAlignment(HorizontalAlignment.CENTER);

        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 11);
        //headerFont.setColor(IndexedColors.RED.getIndex());

        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);

        Font text = workbook.createFont();
        //headerFont.setBold(true);
        text.setFontHeightInPoints((short) 10);
        //headerFont.setColor(IndexedColors.RED.getIndex());

        CellStyle textCell = workbook.createCellStyle();
        textCell.setFont(text);


        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 6));
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 6));
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 0, 6));
        sheet.addMergedRegion(new CellRangeAddress(3, 3, 0, 6));

        Row row = sheet.createRow(0);
        Cell cell1 = row.createCell(0);
        cell1.setCellValue("STARTNE LISTE"); //number of the race
        cell1.setCellStyle(headLineCellStyle);

        row = sheet.createRow(1);
        cell1 = row.createCell(0);
        cell1.setCellValue("KRK Tisin Cvet Senta"); //number of the race
        cell1.setCellStyle(headLineCellStyle);



        row = sheet.createRow(2);
        cell1 = row.createCell(0);
        cell1.setCellValue("Regata \"TISIN CVET\""); //number of the race
        cell1.setCellStyle(headLineCellStyle);



        row = sheet.createRow(3);
        cell1 = row.createCell(0);
        cell1.setCellValue("Senta, 16. jun 2018.godine "); //number of the race
        cell1.setCellStyle(headLineCellStyle);


        // Create a Row
        int i = 5;
        // Create cells

        for (Entry<String, List<String>> entry : treeMap.entrySet()) {

            int num = 1;

            Row headerRow = sheet.createRow(i);
            headerRow.setRowStyle(textCell);

            System.out.println(entry.getKey() + "/" + entry.getValue());
            String[] splitedKey = entry.getKey().split(">");

            Cell cell = headerRow.createCell(0);
            cell.setCellValue(splitedKey[0]); //number of the race
            cell.setCellStyle(headerCellStyle);

            cell = headerRow.createCell(1);
            cell.setCellValue("TRKA");
            cell.setCellStyle(headerCellStyle);

            cell = headerRow.createCell(2);
            cell.setCellValue(splitedKey[1]); //category (k-1, k2, mk1...)
            cell.setCellStyle(headerCellStyle);

            cell = headerRow.createCell(3);
            cell.setCellValue(splitedKey[2]); //age group
            cell.setCellStyle(headerCellStyle);

            cell = headerRow.createCell(5);
            cell.setCellValue("Bodovi");
            cell.setCellStyle(headerCellStyle);


            String time;

            if (Integer.parseInt(splitedKey[0]) < 13) {
                time = "11:";
                Integer temp = 5 * (Integer.parseInt(splitedKey[0]) - 1);
                if (temp < 10) {
                    time += "0" + temp.toString();
                } else {
                    time += temp;
                }
            } else if (Integer.parseInt(splitedKey[0]) < 25) {
                time = "12:";
                Integer temp = 5 * (Integer.parseInt(splitedKey[0]) - 13);
                if (temp < 10) {
                    time += "0" + temp.toString();
                } else {
                    time += temp;
                }
            } else {
                time = "13:";
                Integer temp = 5 * (Integer.parseInt(splitedKey[0]) - 21);
                if (temp < 10) {
                    time += "0" + temp.toString();
                } else {
                    time += temp;
                }
            }

            cell = headerRow.createCell(6);
            cell.setCellValue(time); //start of the race
            cell.setCellStyle(headerCellStyle);

            Collections.shuffle(entry.getValue()); //randomize
            for (String s : entry.getValue()) {
                i++;
                headerRow = sheet.createRow(i);

                cell = headerRow.createCell(1);
                cell.setCellValue(num);
                num++;

                String[] splitValue = s.split(">");

                cell = headerRow.createCell(2);
                cell.setCellValue(splitValue[0].replaceAll("^[\\s\\.\\d]+", "").toUpperCase()); //name

                String[] splitClub = splitValue[1].split("-");

                cell = headerRow.createCell(3);
                cell.setCellValue(splitClub[0]); //club


                cell = headerRow.createCell(4);
                cell.setCellValue(splitClub[1]); //club town


            }


            i += 3;


//                for (int j = 0; j < 7; j++) {
//                    Cell cell = headerRow.createCell(j);
//
//                    if (j = 1)
//
//
//                    cell.setCellValue("adsadasda");
//                    //cell.setCellStyle(headerCellStyle);
//                 }
        }

        for (int k = 0; k < 7; k++) {
            sheet.autoSizeColumn(k);
        }


//        // Resize all columns to fit the content size
//        for(int i = 0; i < columns.length; i++) {
//            sheet.autoSizeColumn(i);
//        }

        // Write the output to a file
        FileOutputStream fileOut = null;
        try {
            fileOut = new FileOutputStream("poi-generated-file.xlsx");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        try {
            workbook.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }
        try {
            fileOut.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

        // Closing the workbook
        try {
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

//        XSSFWorkbook wb = null;
//        wb = new XSSFWorkbook();
//
//        Sheet sheet = wb.createSheet("Tisincvet race");
//
//
//
//        //set sample data
//        for (int i = 0; i < 2; i++) {
//            Row row = sheet.getRow(2 + i);
//            for (int j = 0; j < 2; j++) {
//
//                row.getCell(j).setCellValue("aaaaa");
//            }
//        }
//
//
//
//        // Write the output to a file
//        String file = "out.xlsx";
//
//        FileOutputStream out = null;
//        try {
//            out = new FileOutputStream(file);
//        } catch (FileNotFoundException e) {
//            e.printStackTrace();
//        }
//        try {
//            wb.write(out);
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
//        try {
//            out.close();
//        } catch (IOException e) {
//            e.printStackTrace();
//        }


    }

    /**
     * Create a library of cell styles
     */
    private static Map<String, CellStyle> createStyles(Workbook wb) {
        Map<String, CellStyle> styles = new HashMap<>();
        CellStyle style;
        Font titleFont = wb.createFont();
        titleFont.setFontHeightInPoints((short) 18);
        titleFont.setBold(true);
        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFont(titleFont);
        styles.put("title", style);

        Font monthFont = wb.createFont();
        monthFont.setFontHeightInPoints((short) 11);
        monthFont.setColor(IndexedColors.WHITE.getIndex());
        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFont(monthFont);
        style.setWrapText(true);
        styles.put("header", style);

        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setWrapText(true);
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        styles.put("cell", style);

        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setDataFormat(wb.createDataFormat().getFormat("0.00"));
        styles.put("formula", style);

        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setDataFormat(wb.createDataFormat().getFormat("0.00"));
        styles.put("formula_2", style);

        return styles;
    }

    @Override
    public void start(Stage primaryStage) throws Exception {
        stage = primaryStage;
    }
}
