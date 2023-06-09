import org.hk.dao.WorkWithDB;
import org.hk.services.ReadFromExcel;
import org.hk.services.WriteToExcel;
import org.hk.util.HibernateUtil;

import java.time.LocalDateTime;

import static org.hk.util.Helper.printTime;

public class AppImport {
    public static void main(String[] args) {
        LocalDateTime startLocalDateTime = LocalDateTime.now();

        WorkWithDB.writeRecords(ReadFromExcel.read());
        WriteToExcel.write();

        printTime(startLocalDateTime);
        HibernateUtil.shutdown();
    }
}