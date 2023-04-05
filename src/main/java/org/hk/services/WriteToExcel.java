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
import org.hk.models.OrderReport;
import org.hk.models.QueryParameters;
import org.hk.models.Raw;
import org.hk.models.RecordImport;
import org.hk.util.Helper;
import org.hk.util.HibernateUtil;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Comparator;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import static org.hk.util.Helper.DIR_IMP;
import static org.hk.util.Helper.RAH_201;
import static org.hk.util.Helper.RAH_632;
import static org.hk.util.Helper.deleteFile;

public class WriteToExcel {

    private static final QueryParameters parameters = new QueryParameters();
    private static final DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd.MM.yyyy");
    private static final Set<String> filesToDelete = new HashSet<>();
    private static final String FILE_SEPARATOR = "/";
    private static final String SUFFIX = ".xlsx";
    public static final Map<String, Double> mapWastes = new HashMap<>();

    public static void write() {
        deleteFile(new File(DIR_IMP));
        createAndSaveReport();
    }

    private static void createAndSaveReport() {
        getRecords().parallelStream().filter(r -> r.getProduct() != null).forEach(WriteToExcel::saveReport);
        deleteWrongFiles();
        printMessageAboutMinusValues();
    }

    private static void saveReport(RecordImport recordImport) {
        Session session = HibernateUtil.getSessionFactory().openSession();
        writeReportsToExcelFile(new OrderReport(recordImport, session));
        session.close();
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

    private static File getFileReportToSave(OrderReport report) {
        RecordImport recordImport = report.getRecordImport();
        int year = recordImport.getDate().getYear();
        int monthValue = recordImport.getDate().getMonthValue();
        String folderName = DIR_IMP + FILE_SEPARATOR + year + FILE_SEPARATOR + monthValue;
        createReportFolder(folderName);
        return new File(folderName + FILE_SEPARATOR + recordImport.getOriginDocument() + SUFFIX);
    }

    private static void writeReportsToExcelFile(OrderReport report) {
        report.setFile(getFileReportToSave(report));
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             FileOutputStream fos = new FileOutputStream(report.getFile())) {
            report.setSheet(workbook.createSheet(DIR_IMP));
            insertDataIntoExcelFile(report);
            workbook.write(fos);
            fos.flush();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private static void insertDataIntoExcelFile(OrderReport report) {
        createReportHeader(report);
        addRowsToReport(report);
        createReportFooter(report);
    }

    private static void createReportFolder(String folderName) {
        File folder = new File(folderName);
        if (!folder.exists()) {
            folder.mkdirs();
        }
    }

    private static void createReportHeader(OrderReport report) {
        XSSFSheet sheet = report.getSheet();
        RecordImport recordImport = report.getRecordImport();
        double sum = 0.0;
        int position = report.getRowNumber();
        Row row0 = sheet.createRow(position++);
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

        Row row1 = sheet.createRow(position++);
        Cell cell12 = row1.createCell(2);
        Cell cell14 = row1.createCell(4);
        cell12.setCellValue("Директор");
        cell12.setCellStyle(styleBold);
        cell14.setCellValue("Місько В.І.");
        cell14.setCellStyle(styleBold);
        position++;

        Row row3 = sheet.createRow(position++);
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

        Row row4 = sheet.createRow(position++);
        Cell cell40 = row4.createCell(0);
        cell40.setCellValue(recordImport.getPartner());

        Row row5 = sheet.createRow(position++);
        Cell cell50 = row5.createCell(0);
        cell50.setCellValue(recordImport.getOriginDocument());
        Cell cell53 = row5.createCell(3);
        cell53.setCellValue("Дата переробки");
        position++;

        Row row7 = sheet.createRow(position++);
        Cell cell70 = row7.createCell(0);
        cell70.setCellValue("Бочки");
        Cell cell73 = row7.createCell(3);
        cell73.setCellValue("Комплекти (по входу)");

        List<Raw> rawList = recordImport.getRawList();

        Row row8 = sheet.createRow(position++);
        Cell cell80 = row8.createCell(0);
        cell80.setCellValue(rawList.get(0).getRaw());
        Cell cell85 = row8.createCell(5);
        cell85.setCellValue(rawList.get(0).getCount());
        sum += rawList.get(0).getCount();
        if (rawList.size() > 1) {
            for (int i = 1; i < rawList.size(); i++) {
                sum += addRawInfo(sheet, position++, rawList.get(i));
            }
        }

        Row row100 = sheet.createRow(position++);
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
        cell102.setCellValue("дата");
        cell102.setCellStyle(styleCenter100);
        cell103.setCellValue("Всього вихід");
        cell103.setCellStyle(styleCenter100);
        cell104.setCellValue("м.");
        cell104.setCellStyle(styleCenter100);
        cell105.setCellValue("кг");
        cell105.setCellStyle(styleCenter100);
        report.setSum(sum);
        report.setRowNumber(position);
    }

    private static double addRawInfo(XSSFSheet sheet, int position, Raw raw) {
        Row row = sheet.createRow(position);
        Cell cell0 = row.createCell(0);
        cell0.setCellValue(raw.getRaw());
        Cell cell5 = row.createCell(5);
        cell5.setCellValue(raw.getCount());
        return raw.getCount();
    }

    private static void addRowsToReport(OrderReport report) {
        RecordImport recordImport = report.getRecordImport();
        XSSFSheet sheet = report.getSheet();
        List<RecordImport> reportData = report.getReportData();
        String fileName = report.getFile().getAbsolutePath();
        int position = report.getRowNumber();
        double allCount = 0.0;
        Set<LocalDate> dates = new HashSet<>();
        for (RecordImport data : reportData) {
            Row row = sheet.createRow(position++);
            Cell cell0 = row.createCell(0);
            Cell cell1 = row.createCell(1);
            Cell cell2 = row.createCell(2);
            Cell cell3 = row.createCell(3);
            Cell cell4 = row.createCell(4);
            Cell cell5 = row.createCell(5);

            cell0.setCellValue(data.getProduct());
            cell2.setCellValue(data.getDate().format(formatter));
            cell3.setCellValue(data.getCountResult());
            cell5.setCellValue(data.getCount());

            cell0.setCellStyle(getCellStyle(cell0));
            cell1.setCellStyle(getCellStyle(cell1));
            cell2.setCellStyle(getCellStyle(cell2));
            cell3.setCellStyle(getCellStyle(cell3));
            cell4.setCellStyle(getCellStyle(cell4));
            cell5.setCellStyle(getCellStyle(cell5));

            allCount += data.getCount();
            dates.add(data.getDate());
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
        double waste = report.getSum() - allCount;
        mapWastes.put(report.getFile().getAbsolutePath(), waste);
        cellSumLast.setCellValue(waste);

        if (dates.size() > 0) {
            String first = dates.stream().min(Comparator.naturalOrder()).get().format(formatter);
            String last = dates.stream().max(Comparator.naturalOrder()).get().format(formatter);
            Cell cellDates = row6.createCell(3);
            cellDates.setCellValue(first + "-" + last);
            sheet.addMergedRegion(new CellRangeAddress(6, 6, 3, 4));
        } else {
            filesToDelete.add(fileName);
        }
        report.setRowNumber(++position);
    }

    private static void createReportFooter(OrderReport report) {
        XSSFSheet sheet = report.getSheet();
        int rowNumber = report.getRowNumber();
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

    private static void deleteWrongFiles() {
        System.out.println("------------ Files to delete: ------------------");
        filesToDelete.forEach(System.out::println);
        filesToDelete.parallelStream().map(File::new).forEach(Helper::deleteFile);
    }

    private static void printMessageAboutMinusValues() {
        System.out.println("------ Прихідні накладні в яких відходи від'ємні: -------------");
        mapWastes.entrySet().stream()
                .filter(e -> e.getValue() < 0)
                .forEach(System.out::println);
        System.out.println("---------------------------------------------------------------");
    }
}