package org.hk.services;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.hibernate.Session;
import org.hk.dao.WorkWithDB;
import org.hk.models.QueryParameters;
import org.hk.models.RecordImport;
import org.hk.util.HibernateUtil;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.IntStream;

import static org.hk.util.Helper.DIR_IMP;
import static org.hk.util.Helper.RAH_201;
import static org.hk.util.Helper.RAH_23;
import static org.hk.util.Helper.RAH_25;
import static org.hk.util.Helper.RAH_26;
import static org.hk.util.Helper.RAH_632;
import static org.hk.util.Helper.deleteFile;

public class WriteToExcel {

    private static final List<RecordImport> reportData = new ArrayList<>();
    private static final QueryParameters parameters = new QueryParameters();
    private static final String FILE_SEPARATOR = "/";
    private static final String SUFFIX = ".xlsx";

    public static void write() {
        deleteFile(new File(DIR_IMP));
        createAndSaveReport();
    }

    private static void createAndSaveReport() {
        getRecords().stream().forEach(WriteToExcel::saveReport);
    }

    private static void saveReport(RecordImport recordImport) {
        parameters.setDt(RAH_23);
        parameters.setKt(RAH_201);
        List<RecordImport> records = getRecordsByDocument(recordImport.getCompareDocument());
        createReportData(records);
        writeReportsToExcelFile(recordImport);
        System.out.println(recordImport.getOriginDocument() + " - " + recordImport.getDate());
    }

    private static void createReportData(List<RecordImport> records) {
        reportData.clear();
        records.forEach(record -> {
            parameters.setDt(RAH_23);
            parameters.setKt(RAH_25);
            reportData.addAll(getRecordsByDocument(record.getCompareDocument()));
            parameters.setDt(RAH_201);
            parameters.setKt(RAH_25);
            reportData.addAll(getRecordsByDocument(record.getCompareDocument()));
            setProductNameAndCountResult();
        });
    }

    private static void setProductNameAndCountResult() {
        reportData.forEach(rec -> {
            parameters.setDt(RAH_26);
            parameters.setKt(RAH_23);
            List<RecordImport> names = getRecordsByDocument(rec.getCompareDocument());
            if (names.size() > 0) {
                if (rec.getProduct() == null) {
                    rec.setProduct(names.get(0).getProduct());
                }
                rec.setCountResult(names.get(0).getCount());
            }
        });
    }

    private static List<RecordImport> getRecords() {
        Session session = HibernateUtil.getSessionFactory().openSession();
        parameters.setSession(session);
        parameters.setDt(RAH_201);
        parameters.setKt(RAH_632);
        List<RecordImport> recordsByDtKt = WorkWithDB.getRecordsByDtKt(parameters);
        session.close();
        return recordsByDtKt;
    }

    private static List<RecordImport> getRecordsByDocument(String doc) {
        Session session = HibernateUtil.getSessionFactory().openSession();
        parameters.setSession(session);
        parameters.setDocument(doc);
        List<RecordImport> recordsByDtKt = WorkWithDB.getRecordsByDtKtAndCriteria(parameters);
        session.close();
        return recordsByDtKt;
    }

    private static void writeReportsToExcelFile(RecordImport recordImport) {
        try {
            File file = getFileReportToSave(recordImport);
            FileOutputStream fos = new FileOutputStream(file);
            saveReportToExcel(fos, recordImport);
            fos.flush();
            fos.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static File getFileReportToSave(RecordImport recordImport) {
        int year = recordImport.getDate().getYear();
        int monthValue = recordImport.getDate().getMonthValue();
        String folderName = DIR_IMP + FILE_SEPARATOR + year + FILE_SEPARATOR + monthValue;
        createReportFolder(folderName);
        return new File(folderName + FILE_SEPARATOR + recordImport.getOriginDocument() + SUFFIX);
    }

    private static void saveReportToExcel(FileOutputStream fos, RecordImport recordImport) {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            XSSFSheet sheet = workbook.createSheet(DIR_IMP);
            createReportHeader(sheet, recordImport);
            int rowNumber = addRowsToReport(sheet, recordImport);
            createReportFooter(rowNumber, sheet);
            workbook.write(fos);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private static void createReportFolder(String folderName) {
        File folder = new File(folderName);
        if (!folder.exists()) {
            folder.mkdirs();
        }
    }

    private static void createReportHeader(XSSFSheet sheet, RecordImport recordImport) {
        Row row0 = sheet.createRow(0);
        Cell cell00 = row0.createCell(0);
        Cell cell02 = row0.createCell(2);
        cell00.setCellValue("ТзОВ \"Хінкель-Когут\"");
        cell02.setCellValue("Затверджую");

        CellStyle styleBold = cell00.getSheet().getWorkbook().createCellStyle();
        XSSFFont fontBold = (XSSFFont) cell00.getSheet().getWorkbook().createFont();
        fontBold.setBold(true);
        styleBold.setFont(fontBold);

        cell00.setCellStyle(styleBold);
        cell02.setCellStyle(styleBold);

        Row row1 = sheet.createRow(1);
        Cell cell12 = row1.createCell(2);
        Cell cell14 = row1.createCell(4);
        cell12.setCellValue("Директор");
        cell12.setCellStyle(styleBold);
        cell14.setCellValue("Місько В.І.");
        cell14.setCellStyle(styleBold);

        Row row5 = sheet.createRow(5);
        Cell cell50 = row5.createCell(0);
        cell50.setCellValue("Акт переробки сировини");
        CellStyle styleCenter50 = cell50.getSheet().getWorkbook().createCellStyle();
        setCenterInStyle(styleCenter50);
        XSSFFont fontBold50 = (XSSFFont) cell50.getSheet().getWorkbook().createFont();
        fontBold50.setBold(true);
        fontBold50.setFontHeight(14.0);
        styleCenter50.setFont(fontBold);
        cell50.setCellStyle(styleCenter50);

        sheet.addMergedRegion(new CellRangeAddress(5, 5, 0, 5));

        Row row6 = sheet.createRow(6);
        Cell cell60 = row6.createCell(0);
        cell60.setCellValue(recordImport.getPartner());

        Row row7 = sheet.createRow(7);
        Cell cell70 = row7.createCell(0);
        cell70.setCellValue(recordImport.getOriginDocument());
        Cell cell73 = row7.createCell(3);
        cell73.setCellValue("Дата переробки");

        Row row8 = sheet.createRow(8);
        Cell cell80 = row8.createCell(0);
        cell80.setCellValue("Дата входу");
        Cell cell81 = row8.createCell(1);
        cell81.setCellValue(recordImport.getDate().format(DateTimeFormatter.ofPattern("dd.MM.yyyy")));
        //TODO need add developing dates to cell84

        Row row10 = sheet.createRow(10);
        Cell cell100 = row10.createCell(0);
        cell100.setCellValue("Комплекти (фактично)");
        Cell cell102 = row10.createCell(2);
        cell102.setCellValue("Комплекти (по входу)");
        Cell cell105 = row10.createCell(5);
        cell105.setCellValue(recordImport.getCount());

        Row row120 = sheet.createRow(12);
        Cell cell120 = row120.createCell(0);
        Cell cell121 = row120.createCell(1);
        Cell cell122 = row120.createCell(2);
        Cell cell123 = row120.createCell(3);
        Cell cell124 = row120.createCell(4);
        Cell cell125 = row120.createCell(5);

        CellStyle styleCenter120 = cell70.getSheet().getWorkbook().createCellStyle();
        setCenterInStyle(styleCenter120);
        styleCenter120.setBorderBottom(BorderStyle.MEDIUM);
        styleCenter120.setBorderTop(BorderStyle.MEDIUM);
        styleCenter120.setBorderLeft(BorderStyle.MEDIUM);
        styleCenter120.setBorderRight(BorderStyle.MEDIUM);
        cell120.setCellValue("Калібр");
        cell120.setCellStyle(styleCenter120);
        cell121.setCellValue("довжина, м");
        cell121.setCellStyle(styleCenter120);
        cell122.setCellValue("позначки");
        cell122.setCellStyle(styleCenter120);
        cell123.setCellValue("Всього");
        cell123.setCellStyle(styleCenter120);
        cell124.setCellValue("м.");
        cell124.setCellStyle(styleCenter120);
        cell125.setCellValue("кг");
        cell125.setCellStyle(styleCenter120);
    }

    private static int addRowsToReport(XSSFSheet sheet, RecordImport recordImport) {
//TODO need add developing dates to cell84
        return 13;
    }

    private static void createReportFooter(int rowNumber, XSSFSheet sheet) {
        Row row = sheet.createRow(rowNumber);
        Cell cell0 = row.createCell(0);

        Row rowPrepared = sheet.createRow(rowNumber + 2);
        Cell preparedCell = rowPrepared.createCell(0);
        preparedCell.setCellValue("Заступник директора по виробництву");
        Cell firstPerson = rowPrepared.createCell(4);
        firstPerson.setCellValue("Гладьо Б.М.");

        Row rowReview = sheet.createRow(rowNumber + 4);
        Cell reviewCell = rowReview.createCell(0);
        reviewCell.setCellValue("Головний технолог");

        CellStyle styleBold = preparedCell.getSheet().getWorkbook().createCellStyle();
        XSSFFont fontBold = (XSSFFont) preparedCell.getSheet().getWorkbook().createFont();
        fontBold.setBold(true);
        styleBold.setFont(fontBold);

        preparedCell.setCellStyle(styleBold);
        firstPerson.setCellStyle(styleBold);
        reviewCell.setCellStyle(styleBold);

        IntStream.range(0, 5).forEach(sheet::autoSizeColumn);

        sheet.getPrintSetup().setLandscape(true);
        sheet.setFitToPage(true);
        sheet.getPrintSetup().setFitWidth((short) 1);
        sheet.getPrintSetup().setFitHeight((short) 10);
    }

    private static CellStyle getCellStyle(Cell cellNumberOfRow) {
        CellStyle style = cellNumberOfRow.getSheet().getWorkbook().createCellStyle();
        style.setFont(cellNumberOfRow.getSheet().getWorkbook().createFont());
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    private static void setCenterInStyle(CellStyle style) {
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
    }
}