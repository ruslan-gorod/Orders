package services;

import models.Content;
import models.RecordImport;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.IOException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static util.Helper.CEH;
import static util.Helper.DELIMITER;
import static util.Helper.RAH_201;
import static util.Helper.RAH_23;
import static util.Helper.RAH_25;
import static util.Helper.RAH_26;
import static util.Helper.RAH_632;

public class ReadFromExcel {
    private static final List<RecordImport> records = new ArrayList<>();
    private static final File[] files = new File(".").listFiles();
    private static final Map<String, String> docRecordMap = new HashMap<>();

    public static List<RecordImport> read() {
        assert files != null;
        Arrays.stream(files).forEach(ReadFromExcel::processFile);
        return records;
    }

    public static Map<String, String> getDocRecordMap() {
        return docRecordMap;
    }

    private static void processFile(File file) {
        try {
            Workbook wb = WorkbookFactory.create(file);
            readAndCreateRecords(wb);
            wb.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void readAndCreateRecords(Workbook wb) {
        for (Row r : wb.getSheetAt(0)) {
            RecordImport recordImport = createRecordImport(r);
            if (recordImport.getCount() != 0) {
                records.add(recordImport);
                if (RAH_201.equals(recordImport.getDt()) && RAH_632.equals(recordImport.getKt())) {
                    String key = recordImport.getOriginDocument() + DELIMITER + recordImport.getDate();
                    docRecordMap.put(key, recordImport.getContent().getPartner());
                }
            }
        }
    }

    private static RecordImport createRecordImport(Row r) {
        Content content = getContent(r);
        return RecordImport.builder()
                .date(getRecordLocalDate(r))
                .originDocument(getStringCellValueByPosition(r, 1))
                .compareDocument(content.getDocument())
                .dt(getStringCellValueByPosition(r, 3))
                .kt(getStringCellValueByPosition(r, 4))
                .content(content)
                .count(getCount(r))
                .sum(getRecordSum(r))
                .build();
    }

    private static LocalDate getRecordLocalDate(Row r) {
        return LocalDate.parse(r.getCell(0).getStringCellValue(),
                DateTimeFormatter.ofPattern("dd.MM.yy"));
    }

    private static String getStringCellValueByPosition(Row r, int position) {
        return r.getCell(position).getStringCellValue();
    }

    private static Content getContent(Row r) {
        String value = getStringCellValueByPosition(r, 2);
        String[] recordValue = value.split("\n");
        String doc = getStringCellValueByPosition(r, 1);
        Content content = new Content();
        String dt = getStringCellValueByPosition(r, 3);
        String kt = getStringCellValueByPosition(r, 4);
        if (RAH_201.equals(dt) && RAH_632.equals(kt)) {
            content.setProduct(recordValue[2]);
            content.setDocument(recordValue[3]);
            content.setPartner(recordValue[4]);
        }
        if ((RAH_23.equals(dt) && RAH_201.equals(kt) && CEH.equals(recordValue[1]))
                || (RAH_23.equals(dt) && RAH_25.equals(kt))) {
            String date = " (" + getStringCellValueByPosition(r, 0) + ")";

            content.setDocument(doc.substring(doc.length() - 10) + date);
        }
        if (RAH_201.equals(dt) && RAH_25.equals(kt) && doc.contains("Операция")) {
            content.setProduct(recordValue[2]);
            content.setDocument(recordValue[6]);
        }
        if (RAH_26.equals(dt) && RAH_23.equals(kt) && value.contains("Амортизация")) {
            content.setProduct(recordValue[2]);
            content.setDocument(recordValue[3]);
        }
        return content;
    }

    private static double getRecordSum(Row r) {
        return getaDoubleValueByPosition(r, 5);
    }

    private static double getCount(Row r) {
        return getaDoubleValueByPosition(r, 6);
    }

    private static double getaDoubleValueByPosition(Row r, int position) {
        return r.getCell(position).toString().trim().length() > 0 ?
                r.getCell(position).getNumericCellValue() : 0;
    }
}
