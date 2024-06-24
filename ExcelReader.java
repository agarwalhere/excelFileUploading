/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.myapp;



import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.Vector;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.Cell;
import static org.apache.poi.ss.usermodel.CellType.NUMERIC;
import static org.apache.poi.ss.usermodel.CellType.STRING;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelReader {
    
    private static final Logger LOGGER = Logger.getLogger(ExcelReader.class.getName());
    private static final SimpleDateFormat DATE_FORMAT = new SimpleDateFormat("yyyy/MM/dd");
    private static final int THREAD_POOL_SIZE = 10;
    

    public void processExcelFile(File file) throws InterruptedException {
        LOGGER.log(Level.INFO, "Processing Excel file: {0}", file.getAbsolutePath());
        try {
            Class.forName("com.mysql.cj.jdbc.Driver");
        } catch (ClassNotFoundException ex) {
            Logger.getLogger(ExcelReader.class.getName()).log(Level.SEVERE, null, ex);
        }
        ExecutorService executor = Executors.newFixedThreadPool(THREAD_POOL_SIZE);
        try (FileInputStream fis = new FileInputStream(file);
             XSSFWorkbook workbook = new XSSFWorkbook(fis);
             SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook(workbook,100);
             Connection conn = DriverManager.getConnection("jdbc:mysql://localhost:3306/ihat?autoReconnect=true&maxReconnects=10&socketTimeout=60000", "root", "shourey@2003")) {

            conn.setAutoCommit(false);
            String sql = "INSERT INTO users2 VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?"
                    + ", ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?"
                    + ", ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?"
                    + ", ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?"
                    + ", ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";
            try (PreparedStatement statement = conn.prepareStatement(sql)) {
                Sheet sheet = sxssfWorkbook.getSheetAt(0);
                int batchSize = 100;
                int count = 0;

                for (Row row : sheet) {
                    executor.submit(() -> processRow(row, statement));

                    if (++count % batchSize == 0) {
                        executor.shutdown();
                        executor.awaitTermination(1, TimeUnit.HOURS);
                        statement.executeBatch();
                        conn.commit();
                        LOGGER.log(Level.INFO, "Committed batch at row: {0}", row.getRowNum());
                        sxssfWorkbook.getSheetAt(0).flushRows(100);
                        executor = Executors.newFixedThreadPool(THREAD_POOL_SIZE);
                    }
                }
                executor.shutdown();
                executor.awaitTermination(1, TimeUnit.HOURS);
                statement.executeBatch(); 
                conn.commit();
                LOGGER.log(Level.INFO, "All rows processed and committed successfully.");
            } catch (SQLException e) {
                conn.rollback();
                LOGGER.log(Level.SEVERE, "Failed to process Excel file, transaction rolled back.", e);
            }
        } catch (IOException | SQLException e) {
            LOGGER.log(Level.SEVERE, "Failed to process Excel file", e);
        }
    }
    private String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }
    
    private void processRow(Row row, PreparedStatement statement) {
        try {
            int id = getCellValueAsint(row.getCell(0));
            Vector<String> array1 = new Vector<>(20);
            Vector<Integer> array2 = new Vector<>(129);
            int i = 1;
            for (; i < 21; i++) {
                array1.add(getCellValueAsString(row.getCell(i)));
            }
            for (; i < 150; i++) {
                array2.add(getCellValueAsint(row.getCell(i)));
            }
            statement.setInt(1, id);
            int k = 2, a = 0;
            for (; k < 22; k++) {
                statement.setString(k, array1.get(a));
                a++;
            }
            a = 0;
            for (; k < 151; k++) {
                statement.setInt(k, array2.get(a));
                a++;
            }
            statement.addBatch();
        } catch (SQLException e) {
            LOGGER.log(Level.SEVERE, "Failed to process row", e);
        }
    }
    
  


  
    private int getCellValueAsint(Cell cell) {
        if(cell == null){
            return 0;
        }
        switch(cell.getCellType()){
            case NUMERIC:
                return (int)cell.getNumericCellValue();
            case STRING:
                return Integer.parseInt(cell.getStringCellValue());
            default:
                return 0;
        }
    }

    
    
}
