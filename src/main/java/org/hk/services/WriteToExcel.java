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
import org.hk.models.Raw;
import org.hk.models.RecordImport;
import org.hk.util.HibernateUtil;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import static org.hk.util.Helper.DIR_IMP;
import static org.hk.util.Helper.RAH_201;
import static org.hk.util.Helper.RAH_23;
import static org.hk.util.Helper.RAH_25;
import static org.hk.util.Helper.RAH_26;
import static org.hk.util.Helper.RAH_632;
import static org.hk.util.Helper.deleteFile;

public class WriteToExcel {

    private static final List<RecordImport> reportData = new ArrayList<>();
    private static final Map<String, RecordImport> writtenRecords = new HashMap<>();
    private static final QueryParameters parameters = new QueryParameters();
    private static final DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd.MM.yyyy");
    private static final String FILE_SEPARATOR = "/";
    private static final String DASH = " - ";
    private static final String SUFFIX = ".xlsx";

    public static void write() {
        deleteFile(new File(DIR_IMP));
        createAndSaveReport();
    }

    private static void createAndSaveReport() {
        getRecords().stream().filter(r -> r.getProduct() != null).forEach(WriteToExcel::saveReport);
    }

    private static void saveReport(RecordImport recordImport) {
        parameters.setDt(RAH_23);
        parameters.setKt(RAH_201);
        String document = recordImport.getCompareDocument();
        List<RecordImport> records = getRecordsByDocument(document);
        createReportData(records);

        RecordImport writtenRecord = writtenRecords.get(document);
        recordImport.getRawList().add(new Raw(recordImport.getProduct(), recordImport.getCount()));
        if (writtenRecord == null) {
            writtenRecords.put(document, recordImport);
        } else {
            recordImport.getRawList().add(new Raw(writtenRecord.getProduct(), writtenRecord.getCount()));
        }
        writeReportsToExcelFile(recordImport);
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
            double sum = createReportHeader(sheet, recordImport);
            int rowNumber = addRowsToReport(sheet, recordImport, sum);
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

    private static double createReportHeader(XSSFSheet sheet, RecordImport recordImport) {
        double sum = 0.0;
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

        Row row3 = sheet.createRow(3);
        Cell cell30 = row3.createCell(0);
        cell30.setCellValue("Акт переробки сировини");
        CellStyle styleCenter30 = cell30.getSheet().getWorkbook().createCellStyle();
        setCenterInStyle(styleCenter30);
        XSSFFont fontBold30 = (XSSFFont) cell30.getSheet().getWorkbook().createFont();
        fontBold30.setBold(true);
        fontBold30.setFontHeight(16.0);
        styleCenter30.setFont(fontBold30);
        cell30.setCellStyle(styleCenter30);

        sheet.addMergedRegion(new CellRangeAddress(3, 3, 0, 5));

        Row row4 = sheet.createRow(4);
        Cell cell40 = row4.createCell(0);
        cell40.setCellValue(recordImport.getPartner());

        Row row5 = sheet.createRow(5);
        Cell cell50 = row5.createCell(0);
        cell50.setCellValue(recordImport.getOriginDocument());
        Cell cell53 = row5.createCell(3);
        cell53.setCellValue("Дата переробки");

        Row row7 = sheet.createRow(7);
        Cell cell73 = row7.createCell(3);
        cell73.setCellValue("Комплекти (по входу)");

        List<Raw> rawList = recordImport.getRawList();

        Row row8 = sheet.createRow(8);
        Cell cell80 = row8.createCell(0);
        cell80.setCellValue(rawList.get(0).getRaw());
        Cell cell85 = row8.createCell(5);
        cell85.setCellValue(rawList.get(0).getCount());
        sum += rawList.get(0).getCount();
        if (rawList.size() > 1) {
            Row row9 = sheet.createRow(9);
            Cell cell90 = row9.createCell(0);
            cell90.setCellValue(rawList.get(1).getRaw());
            Cell cell95 = row9.createCell(5);
            cell95.setCellValue(rawList.get(1).getCount());
            sum += rawList.get(1).getCount();
        }

        Row row100 = sheet.createRow(10);
        Cell cell100 = row100.createCell(0);
        Cell cell101 = row100.createCell(1);
        Cell cell102 = row100.createCell(2);
        Cell cell103 = row100.createCell(3);
        Cell cell104 = row100.createCell(4);
        Cell cell105 = row100.createCell(5);

        CellStyle styleCenter100 = cell50.getSheet().getWorkbook().createCellStyle();
        setCenterInStyle(styleCenter100);
        styleCenter100.setBorderBottom(BorderStyle.MEDIUM);
        styleCenter100.setBorderTop(BorderStyle.MEDIUM);
        styleCenter100.setBorderLeft(BorderStyle.MEDIUM);
        styleCenter100.setBorderRight(BorderStyle.MEDIUM);
        cell100.setCellValue("Калібр");
        cell100.setCellStyle(styleCenter100);
        cell101.setCellValue("довжина, м");
        cell101.setCellStyle(styleCenter100);
        cell102.setCellValue("позначки");
        cell102.setCellStyle(styleCenter100);
        cell103.setCellValue("Всього вихід");
        cell103.setCellStyle(styleCenter100);
        cell104.setCellValue("м.");
        cell104.setCellStyle(styleCenter100);
        cell105.setCellValue("кг");
        cell105.setCellStyle(styleCenter100);
        return sum;
    }

    private static int addRowsToReport(XSSFSheet sheet, RecordImport recordImport, double sum) {
        int position = 11;
        double allCount = 0.0;
        Set<LocalDate> dates = new HashSet<>();
        String noInfo = recordImport.getOriginDocument() + DASH
                + recordImport.getDate() + DASH + recordImport.getProduct();
        for (RecordImport report : reportData) {
            Row row = sheet.createRow(position++);
            Cell cell0 = row.createCell(0);
            Cell cell1 = row.createCell(1);
            Cell cell2 = row.createCell(2);
            Cell cell3 = row.createCell(3);
            Cell cell4 = row.createCell(4);
            Cell cell5 = row.createCell(5);

            cell0.setCellValue(report.getProduct());
            cell3.setCellValue(report.getCountResult());
            cell5.setCellValue(report.getCount());

            cell0.setCellStyle(getCellStyle(cell0));
            cell1.setCellStyle(getCellStyle(cell1));
            cell2.setCellStyle(getCellStyle(cell2));
            cell3.setCellStyle(getCellStyle(cell3));
            cell4.setCellStyle(getCellStyle(cell4));
            cell5.setCellStyle(getCellStyle(cell5));

            allCount += report.getCount();
            dates.add(report.getDate());
            noInfo = report.getOriginDocument() + DASH + report.getDate();
        }
        Row row = sheet.createRow(position++);
        Cell cell0 = row.createCell(0);
        Cell cell1 = row.createCell(1);
        Cell cell2 = row.createCell(2);
        Cell cell3 = row.createCell(3);
        Cell cell4 = row.createCell(4);
        Cell cell5 = row.createCell(5);

        cell0.setCellValue("ВСЬОГО");
        cell5.setCellValue(allCount);

        cell0.setCellStyle(getCellStyle(cell0));
        cell1.setCellStyle(getCellStyle(cell1));
        cell2.setCellStyle(getCellStyle(cell2));
        cell3.setCellStyle(getCellStyle(cell3));
        cell4.setCellStyle(getCellStyle(cell4));
        cell5.setCellStyle(getCellStyle(cell5));

        Row row6 = sheet.createRow(6);
        Cell cell60 = row6.createCell(0);
        cell60.setCellValue("Дата входу");
        Cell cell61 = row6.createCell(1);
        cell61.setCellValue(recordImport.getDate().format(formatter));

        Row rowLast = sheet.createRow(1 + position);
        Cell cellLast = rowLast.createCell(0);
        cellLast.setCellValue("Технологічні відходи в т.ч. сіль");
        Cell cellSumLast = rowLast.createCell(5);
        cellSumLast.setCellValue(sum - allCount);

        if (dates.size() > 0) {
            String first = dates.stream().min(Comparator.naturalOrder()).get().format(formatter);
            String last = dates.stream().max(Comparator.naturalOrder()).get().format(formatter);
            Cell cellDates = row6.createCell(3);
            cellDates.setCellValue(first + "-" + last);
            sheet.addMergedRegion(new CellRangeAddress(6, 6, 3, 4));
        } else {
            System.out.println("no data: " + noInfo);
        }
        return ++position;
    }

    private static void createReportFooter(int rowNumber, XSSFSheet sheet) {
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

        sheet.setColumnWidth(0, 9960);
        sheet.setColumnWidth(1, 3140);
        sheet.setColumnWidth(2, 3140);
        sheet.setColumnWidth(3, 3430);
        sheet.setColumnWidth(4, 1450);
        sheet.setColumnWidth(5, 2280);

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