import dao.WorkWithDB;
import models.RecordImport;
import services.ReadFromExcel;

import java.time.Duration;
import java.time.LocalDateTime;
import java.util.List;
import java.util.concurrent.TimeUnit;

public class AppImport {
    public static void main(String[] args) {
        LocalDateTime startLocalDateTime = LocalDateTime.now();

        List<RecordImport> records = ReadFromExcel.read();
        System.out.println(ReadFromExcel.getDocRecordMap().size());
        WorkWithDB.writeRecords(records);

        System.out.println("Completed");

        printTime(startLocalDateTime);
        System.exit(0);
    }

    private static void printTime(LocalDateTime startLocalDateTime) {
        long millis = Duration.between(startLocalDateTime, LocalDateTime.now()).toMillis();
        long minutes = TimeUnit.MILLISECONDS.toMinutes(millis);
        String time = String.format("%d minutes %d seconds", minutes,
                TimeUnit.MILLISECONDS.toSeconds(millis) - TimeUnit.MINUTES.toSeconds(minutes));
        System.out.printf("Time taken: %s%n", time);
    }
}
