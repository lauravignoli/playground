package genTests;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.joda.time.DateTime;

public class BsisImportTestSpreadsheetGenerator {

  private static final int ROWS = 20000;
  private static final int DONOR_ID_OFFSET = 0;
  private static final int DONATION_NR_OFFSET = 1000000;

  public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException {
    HSSFWorkbook workbook = new HSSFWorkbook();
    createLocations(workbook);
    createDonors(workbook);
    createDonations(workbook);
    createDeferrals(workbook);
    createOutcomes(workbook);

    try {
      FileOutputStream out = new FileOutputStream(new File("/home/laura/lauraTest.xlsx"));
      workbook.write(out);
      out.close();
      System.out.println("Excel written successfully..");

    } catch (FileNotFoundException e) {
      e.printStackTrace();
    } catch (IOException e) {
      e.printStackTrace();
    }
  }

  private static void createLocations(HSSFWorkbook workbook) {
    HSSFSheet sheet = workbook.createSheet("Locations");

    Row firstRow = sheet.createRow(0);
    // Set column names
    String[] columnNames = {"name", "isUsageSite", "isMobileSite", "isVenue", "isDeleted", "notes"};
    populateCells(firstRow, columnNames);

    for (int rowNum = 1; rowNum < 500; rowNum++) {
      Row row = sheet.createRow(rowNum);
      Object[] strings = {"loc" + rowNum, false, false, true, false, ""};
      populateCells(row, strings);
    }
  }

  private static void createDonors(HSSFWorkbook workbook) {
    HSSFSheet sheet = workbook.createSheet("Donors");

    Row firstRow = sheet.createRow(0);
    // Set column names
    String[] columnNames = {"externalDonorId", "title", "firstName", "middleName", "lastName", "callingName", "gender",
        "preferredLanguage", "birthDate", "bloodAbo", "bloodRh", "notes", "venue", "idType", "idNumber",
        "preferredContactType", "mobileNumber", "homeNumber", "workNumber", "email", "preferredAddressType",
        "homeAddressLine1", "homeAddressLine2", "homeAddressCity", "homeAddressProvince", "homeAddressDistrict",
        "homeAddressState", "homeAddressCountry", "homeAddressZipcode", "workAddressLine1", "workAddressLine2",
        "workAddressCity", "workAddressProvince", "workAddressDistrict", "workAddressCountry", "workAddressState",
        "workAddressZipcode", "postalAddressLine1", "postalAddressLine2", "postalAddressCity", "postalAddressProvince",
        "postalAddressDistrict", "postalAddressCountry", "postalAddressState", "postalAddressZipcode"};
    populateCells(firstRow, columnNames);

    DateTime birthDate = new DateTime().minusYears(90);
    for (int rowNum = 1; rowNum < ROWS; rowNum++) {
      Row row = sheet.createRow(rowNum);
      birthDate = birthDate.plusMinutes(5);
      Object[] strings =
          {rowNum + DONOR_ID_OFFSET + "", "", "Donor", "Juno", "Lacio", "Lau", "female", "English", birthDate, "A", "+", "Notes", "fff",
              "Passport Number", "", "Email", "27825540216", "217830444", "", "xxx@gmail.xxx", "Work Address",
              "ddd", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "7975"};
      populateCells(row, strings);
    }
  }

  private static void createDonations(HSSFWorkbook workbook) {
    HSSFSheet sheet = workbook.createSheet("Donations");

    Row firstRow = sheet.createRow(0);
    // Set column names
    String[] columnNames = {"externalDonorId", "donationIdentificationNumber", "venue", "donationType", "packType",
        "donationDate", "bleedStartTime", "bleedEndTime", "donorWeight", "bloodPressureSystolic",
        "bloodPressureDiastolic", "donorPulse", "haemoglobinCount", "haemoglobinLevel", "adverseEventType",
        "adverseEventComment", "bloodAbo", "bloodRh", "notes"};
    populateCells(firstRow, columnNames);

    DateTime donationDate = new DateTime().minusYears(90);
    for (int rowNum = 1; rowNum < ROWS; rowNum++) {
      Row row = sheet.createRow(rowNum);
      
      donationDate = donationDate.plusMinutes(6);
      Object[] strings = {rowNum + DONOR_ID_OFFSET + "", DONATION_NR_OFFSET + rowNum + "", "fff", "Voluntary", "Double",
          donationDate,
          donationDate, donationDate, 200, 110, 60, 60, 20, "", "Nausea", "678", "A", "+", "Hello"};
      populateCells(row, strings);
    }
  }

  private static void createDeferrals(HSSFWorkbook workbook) {
    HSSFSheet sheet = workbook.createSheet("Deferrals");

    Row firstRow = sheet.createRow(0);
    // Set column names
    String[] columnNames =
        {"externalDonorId", "venue", "deferralReason", "deferralReasonText", "createdDate", "deferredUntil"};
    populateCells(firstRow, columnNames);

    DateTime donationDate = new DateTime().minusYears(2);
    for (int rowNum = 1; rowNum < 100; rowNum++) {
      Row row = sheet.createRow(rowNum);

      donationDate = donationDate.plusMinutes(30);
      Object[] strings =
          {rowNum + DONOR_ID_OFFSET + "", "fff", "Low weight", "Deferral reason text", donationDate, donationDate.plusYears(20)};
      populateCells(row, strings);
    }
  }

  private static void createOutcomes(HSSFWorkbook workbook) {
    HSSFSheet sheet = workbook.createSheet("Outcomes");

    Row firstRow = sheet.createRow(0);
    // Set column names
    String[] columnNames = {"donationIdentificationNumber", "testedOn", "bloodTestName", "outcome"};
    populateCells(firstRow, columnNames);

    DateTime donationDate = new DateTime().minusYears(2);

    for (int rowNum = 1; rowNum < ROWS; rowNum++) {
      Row row = sheet.createRow(rowNum);

      donationDate = donationDate.plusMinutes(30);
      Object[] strings = {DONATION_NR_OFFSET + rowNum + "", donationDate, "HIV", "POS"};
      populateCells(row, strings);
    }
  }

  private static void populateCells(Row row, Object[] values) {
    int i = 0;
    for (Object value : values) {
      if (value instanceof String) {
        String valueStr = (String) value;
        row.createCell(i).setCellValue(valueStr);
      } else if (value instanceof DateTime) {
        DateTime date = (DateTime) value;
        Cell cell = row.createCell(i);
        cell.setCellValue(date.toDate());

      } else if (value instanceof Boolean) {
        Boolean bool = (Boolean) value;
        row.createCell(i).setCellValue(bool);
      } else if (value instanceof Integer) {
        Integer integer = (Integer) value;
        row.createCell(i).setCellValue(integer);
      }
      i++;
    }
  }
}





