package org.example;

import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.Iterator;
import java.util.concurrent.atomic.AtomicInteger;


public class App {
    public static void main(String[] args) throws IOException {
        String excelFilePath = "Letter (2) (2).xlsx";
        FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
        Workbook workbook = WorkbookFactory.create(inputStream);

        Sheet firstSheet = workbook.getSheetAt(0);
        AtomicInteger count = new AtomicInteger();
        AtomicInteger fileCount = new AtomicInteger();

        PrintWriter writer = new PrintWriter("deleteFilesLetter_" + fileCount.get() + ".txt", "UTF-8");

        writer.println("set context person creator;"); // adding this at the start of each file
        writer.println("start transaction;"); // starting the transaction

        Iterator<Row> iterator = firstSheet.iterator();
        iterator.next(); // skipping the header row

        while (iterator.hasNext()) {
            Row nextRow = iterator.next();
            Cell typeCell = nextRow.getCell(0);
            Cell nameCell = nextRow.getCell(1);
            Cell revisionCell = nextRow.getCell(2);
            Cell idsCell = nextRow.getCell(5);

            if (typeCell != null && nameCell != null && revisionCell != null) {
                writer.println("del bus " + typeCell.toString().trim() + " " + nameCell.toString().trim() + " " + (int) revisionCell.getNumericCellValue() + " format generic file all;");
            }

            if (idsCell != null) {
                String[] ids = idsCell.toString().split(",");
                for (String id : ids) {
                    if (id.trim().startsWith("38435")) {
                        writer.println("del bus " + id.trim() + ";");
                    }
                }
            }

            if (count.incrementAndGet() % 500 == 0) {
                writer.println("commit transaction;"); // committing the transaction
                writer.close();
                writer = new PrintWriter("deleteFilesLetter_" + fileCount.incrementAndGet() + ".txt", "UTF-8");
                writer.println("set context person creator;"); // adding this at the start of each new file
                writer.println("start transaction;"); // starting the transaction
            }
        }

        writer.println("commit transaction;"); // committing the transaction
        writer.close();
        workbook.close();
        inputStream.close();
    }
}
