package org.hk.dao;

import org.hibernate.Session;
import org.hibernate.Transaction;
import org.hibernate.query.Query;
import org.hk.models.QueryParameters;
import org.hk.models.RecordImport;
import org.hk.util.HibernateUtil;

import java.util.List;

public class WorkWithDB {

    public static void writeRecords(List<RecordImport> records) {
        records.forEach(WorkWithDB::saveRecord);
    }

    private static void saveRecord(RecordImport record) {
        Transaction transaction = null;
        try (Session session = HibernateUtil.getSessionFactory().openSession()) {
            transaction = session.beginTransaction();
            session.save(record);
            transaction.commit();
        } catch (Exception e) {
            if (transaction != null) {
                transaction.rollback();
            }
            e.printStackTrace();
        }
    }

    public static List<RecordImport> getRecordsByDtKt(QueryParameters parameters) {
        Query query = parameters.getSession().createQuery(
                "from RecordImport where dt = :dt and kt = :kt order by date asc");
        query.setParameter("dt", parameters.getDt());
        query.setParameter("kt", parameters.getKt());
        return query.list();
    }

    public static List<RecordImport> getRecordsByDtKtAndCriteria(QueryParameters parameters) {
        Query query = parameters.getSession().createQuery(
                "from RecordImport " +
                        "where dt = :dt and kt = :kt and criteriaDocument = : doc " +
                        "order by date asc");
        query.setParameter("dt", parameters.getDt());
        query.setParameter("kt", parameters.getKt());
        query.setParameter("doc", parameters.getDocument());
        return query.list();
    }
}