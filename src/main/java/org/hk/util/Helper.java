package org.hk.util;

import java.io.File;
import java.time.Duration;
import java.time.LocalDateTime;
import java.util.Arrays;
import java.util.Objects;
import java.util.concurrent.TimeUnit;

public class Helper {
    public static final String DIR_IMP = "ordersIMP";
    public static final String CEH = "цех замочування";
    public static final String RAH_201 = "201";
    public static final String RAH_23 = "23";
    public static final String RAH_25 = "25";
    public static final String RAH_26 = "26";
    public static final String RAH_632 = "632";

    public static void deleteFile(File element) {
        if (element.exists() && element.isDirectory()) {
            Arrays.stream(Objects.requireNonNull(element.listFiles())).forEach(Helper::deleteFile);
        }
        element.delete();
    }

    public static void printTime(LocalDateTime startLocalDateTime) {
        long millis = Duration.between(startLocalDateTime, LocalDateTime.now()).toMillis();
        long minutes = TimeUnit.MILLISECONDS.toMinutes(millis);
        String time = String.format("%d minutes %d seconds", minutes,
                TimeUnit.MILLISECONDS.toSeconds(millis) - TimeUnit.MINUTES.toSeconds(minutes));
        System.out.printf("Completed. Time taken: %s%n", time);
    }
}