package org.example;

import org.apache.hc.client5.http.classic.methods.HttpGet;
import org.apache.hc.client5.http.impl.classic.CloseableHttpClient;
import org.apache.hc.client5.http.impl.classic.CloseableHttpResponse;
import org.apache.hc.client5.http.impl.classic.HttpClients;
import org.apache.hc.core5.http.HttpEntity;
import org.apache.hc.core5.http.ParseException;
import org.apache.hc.core5.http.io.entity.EntityUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.net.URLDecoder;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.Iterator;

public class Main {
    public static void main(String[] args) {
        //you need to change this path
        String excelFilePath = "C:\\Users\\mohamed.akram\\Desktop\\date.xlsx"; // Replace with your Excel file path
        try {
            // Open the Excel file
            FileInputStream fileInputStream = new FileInputStream(new File(excelFilePath));
            Workbook workbook = new XSSFWorkbook(fileInputStream);

            // Iterate through rows and cells directly from the workbook
            Iterator<Row> rowIterator = workbook.getSheetAt(0).iterator();

            // Skip the header row
            if (rowIterator.hasNext()) {
                rowIterator.next();
            }

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();

                // Read fromDate and toDate from the first two cells in each row
                Cell fromDateCell = row.getCell(0);
                Cell toDateCell = row.getCell(1);

                if (fromDateCell == null || toDateCell == null) {
                    continue; // Skip the row if any of the date cells are empty
                }

                String fromDate = fromDateCell.getStringCellValue();
                String toDate = toDateCell.getStringCellValue();

                // Create API URL with dates
                String apiUrl = String.format("https://www.mubasher.info/api/1/analysis/stock-statistics/winners?from=%s&market=EGX&to=%s", fromDate, toDate);

                // Make the API request and get the response
                String response = makeApiRequest(apiUrl);

                // Save response to file with the date as the filename
                saveResponseToFile(response, convertDate(fromDate) + ".json");
            }

            workbook.close();
            fileInputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // Method to make API request and return the response as a String
    private static String makeApiRequest(String url) throws IOException, ParseException {
        try (CloseableHttpClient httpClient = HttpClients.createDefault()) {
            HttpGet request = new HttpGet(url);
            try (CloseableHttpResponse response = httpClient.execute(request)) {
                HttpEntity entity = response.getEntity();
                return EntityUtils.toString(entity);
            }
        }
    }

    // Method to save API response to a file
    private static void saveResponseToFile(String response, String fileName) throws IOException {
        //change to the path that you want to save the file
        try (FileWriter fileWriter = new FileWriter(new File(System.getProperty("user.home") + "/Desktop/" + fileName))) {
            fileWriter.write(response);
        }

    }

    private static String convertDate(String urlEncodedDate) {
        try {
            // Decode the URL-encoded date
            String decodedDate = URLDecoder.decode(urlEncodedDate, "UTF-8");

            // Define the input and output date formatters
            DateTimeFormatter inputFormatter = DateTimeFormatter.ofPattern("yyyy/MM/dd");
            DateTimeFormatter outputFormatter = DateTimeFormatter.ofPattern("d-M-yyyy");

            // Parse the decoded date
            LocalDate date = LocalDate.parse(decodedDate, inputFormatter);

            // Format the date to the desired output format
            return date.format(outputFormatter);

        } catch (UnsupportedEncodingException | DateTimeParseException e) {
            // Handle exceptions
            e.printStackTrace();
            return null;
        }
    }
}
