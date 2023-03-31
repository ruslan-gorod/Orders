package org.hk.models;

import lombok.Data;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.hibernate.Session;
import org.hk.dao.WorkWithDB;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

import static org.hk.util.Helper.*;

@Data
public class OrderReport {
    private File file;
    private double sum;
    private int rowNumber = 0;
    private XSSFSheet sheet;
    private final Session session;
    private RecordImport recordImport;
    private List<RecordImport> reportData = new ArrayList<>();
    private QueryParameters parameters = new QueryParameters();
    private static final Map<String, RecordImport> writtenRecords = new ConcurrentHashMap<>();

    public OrderReport(RecordImport recordImport, Session session) {
        this.recordImport = recordImport;
        this.session = session;
        createReportData();
    }

    private void createReportData() {
        parameters.setDt(RAH_23);
        parameters.setKt(RAH_201);
        getRecordsByDocument(recordImport.getCompareDocument())
                .stream()
                .map(RecordImport::getCompareDocument)
                .forEach(document -> {
                    parameters.setDt(RAH_23);
                    parameters.setKt(RAH_25);
                    reportData.addAll(getRecordsByDocument(document));
                    parameters.setDt(RAH_201);
                    parameters.setKt(RAH_25);
                    reportData.addAll(getRecordsByDocument(document));
                    setProductNameAndCountResult(reportData);
                });
        countWrittenRecords();
    }

    private void countWrittenRecords() {
        String document = recordImport.getCompareDocument();
        RecordImport writtenRecord = writtenRecords.get(document);
        recordImport.getRawList().add(new Raw(recordImport));
        if (writtenRecord == null) {
            writtenRecords.put(document, recordImport);
        } else {
            recordImport.getRawList().add(new Raw(writtenRecord));
        }
    }

    private void setProductNameAndCountResult(List<RecordImport> reportData) {
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

    private List<RecordImport> getRecordsByDocument(String doc) {
        parameters.setSession(session);
        parameters.setDocument(doc);
        return WorkWithDB.getRecordsByDtKtAndCriteria(parameters);
    }
}