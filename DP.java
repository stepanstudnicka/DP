package DP;
  import java.io.IOException;
  import java.io.InputStream;
  import java.io.FileInputStream;
  import java.text.ParseException;
  import java.text.SimpleDateFormat;
  import java.util.*;
  import org.apache.poi.ss.usermodel.*;
  import org.apache.poi.xssf.usermodel.XSSFWorkbook;
  import org.apache.poi.ss.usermodel.Row;
  import java.io.FileOutputStream;
   public class DP {
      private static Map<String, FlightInfo> standReservations;
      private static Date systemDate;
      private static String routing;
      private static Map<String, List<String>> standRestrictions = new HashMap<>();


      private static class FlightInfo {
          String tripNumber;
          String scheduledDate;
          String scheduledDate2;
          String scheduledTime;
          String scheduledTime4;
          String acType;
          String acType4;
          String aircraftCategory;
          String timeDifference;
          String reservedStand;
          String airportCode;
          String routing;
          String schengenStatus;
          String airlineCode;
          String scheduledDate3;

          FlightInfo(String tripNumber, String scheduledDate, String scheduledDate2, String scheduledTime,
                     String scheduledTime4, String acType, String acType4, String aircraftCategory,
                     String timeDifference, String reservedStand, String airportCode, String routing, String schengenStatus, String airlineCode, String scheduledDate3) {
              this.tripNumber = tripNumber;
              this.scheduledDate = scheduledDate;
              this.scheduledDate2 = scheduledDate2;
              this.scheduledTime = scheduledTime;
              this.scheduledTime4 = scheduledTime4;
              this.acType = acType;
              this.acType4 = acType4;
              this.aircraftCategory = aircraftCategory;
              this.timeDifference = timeDifference;
              this.reservedStand = reservedStand;
              this.airportCode = airportCode;
              this.routing = routing;
              this.schengenStatus = schengenStatus;
              this.airlineCode = airlineCode; // Initialize the new field
              this.scheduledDate3 = scheduledDate3;

          }
      }
      public static void main(String[] args) {
          try {
              List<String> airportCodes = readLetovyradcasData("letovyradcas.xlsx");
              Map<String, String> airportData = readACallData("ACall.xlsx");
              Map<String, String> flightCodes = readFlightCodes("ACall.xlsx");
              String flightCode = flightCodes.get(routing);
              Set<String> excludedStands = new HashSet<>();
              Map<String, List<String>> airlinePreferences = readAirlinePreferences("preferences.xlsx");
              standRestrictions = readStandRestrictions("restrictions.xlsx");


              InputStream inputStream = Save7.class.getClassLoader().getResourceAsStream("letovyradcas.xlsx");
              if (inputStream != null) {
                  Workbook workbook = new XSSFWorkbook(inputStream);
                  Sheet sheet = workbook.getSheetAt(0);
                  System.out.println("\nWorking with sheet: " + sheet.getSheetName());
                  Scanner scanner = new Scanner(System.in);
                  System.out.print("Enter the desired date (dd.MM.yyyy): ");
                  String desiredDate = scanner.nextLine();
                  standReservations = new HashMap<>();
                  SimpleDateFormat dateFormat = new SimpleDateFormat("dd.MM.yyyy");
                  try {
                      systemDate = dateFormat.parse(desiredDate);
                  } catch (ParseException e) {
                      e.printStackTrace();
                      throw new RuntimeException("Error parsing date", e);
                  }
                  String startDate = findStartDate(sheet, desiredDate);
                  System.out.println("Desired Date: " + desiredDate);
                  System.out.println("Start Date: " + startDate);
                  int matchCount = 0; // Counter for matches

                  if (startDate != null) {
                      for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                          Row rowTripNumber4 = sheet.getRow(i);
                          Cell cellTripNumber4 = rowTripNumber4.getCell(8);
                          if (cellTripNumber4 != null) {
                              cellTripNumber4.setCellType(CellType.STRING);
                              String tripNumber4 = cellTripNumber4.getStringCellValue().trim();
                      	    //System.out.println("Match found for Trip Number4: " + tripNumber4 + " " + i);

                              for (int j = i + 1; j <= sheet.getLastRowNum(); j++) {
                                  Row rowTripNumber = sheet.getRow(j);
                                  Cell cellTripNumber = rowTripNumber.getCell(3);
                                  if (cellTripNumber != null) {
                                      cellTripNumber.setCellType(CellType.STRING);
                                      String tripNumber = cellTripNumber.getStringCellValue().trim();
                              	    //System.out.println("Match found for Trip Number: " + tripNumber + " " + j);

                                      if (tripNumber4.equalsIgnoreCase(tripNumber))  {

                                    	  Cell cellScheduledTime = rowTripNumber.getCell(1);
                                          Cell cellScheduledTime4 = rowTripNumber4.getCell(1);
                                          Cell cellScheduledTime6 = rowTripNumber4.getCell(1);

                                          String scheduledTime = getStringValue(cellScheduledTime);
                                          String scheduledTime4 = getStringValue(cellScheduledTime4);
                                          String scheduledTime6 = getStringValue(cellScheduledTime6);

                                          Cell cellRouting = rowTripNumber.getCell(12);
                                          String routing = getStringValue(cellRouting);
                                          String schengenStatus = getSchengenStatus(routing, airportData);
                                          Cell cellScheduledDate = rowTripNumber.getCell(0);
                                          Cell cellScheduledDate2 = rowTripNumber4.getCell(6);
                                          //Cell cellScheduledTime = rowTripNumber.getCell(1);
                                          //Cell cellScheduledTime4 = rowTripNumber4.getCell(1);
                                          Cell cellACType = rowTripNumber.getCell(11);
                                          Cell cellACType4 = rowTripNumber4.getCell(11);
                                          Cell cellWingSpan = rowTripNumber.getCell(1);
                                          Cell cellAirlineCode = rowTripNumber.getCell(2); // Assuming airline code is in column C
                                          String airlineCode = getStringValue(cellAirlineCode); // Use existing getStringValue method
                                          String scheduledDate = getStringValue(cellScheduledDate);
                                          String scheduledDate2 = getStringValue(cellScheduledDate2);
                                          //String scheduledTime = getStringValue(cellScheduledTime);
                                          //String scheduledTime4 = getStringValue(cellScheduledTime4);
                                          String scheduledDate3 = getStringValue(rowTripNumber4.getCell(0)); // extracted from matching row
                                          String scheduledDate4 = getStringValue(rowTripNumber.getCell(6)); // extracted from base row

                                         
                                          
                                          if ((scheduledDate.equals(desiredDate) || scheduledDate2.equals(desiredDate)) && // Either date matches desired date
                                        		    ((scheduledDate.equals(desiredDate) && scheduledDate2.equals(desiredDate)) || // Both dates match desired date
                                        		    (scheduledDate2.equals(desiredDate) && scheduledDate.equals(desiredDate)))) { // Both dates match desired date
                                        	    System.out.println("Debug: Date matches found for " + scheduledDate + " or " + scheduledDate2 + " " + tripNumber + " " + tripNumber4);

                                        	    
                                        	  String acType = getACType(cellACType);
                                              String acType4 = getACType(cellACType4);
                                              double wingSpan = getNumericValue(cellWingSpan);
                                              String aircraftCategory = getCategoryFromAllAC(acType, "allAC.xlsx");
                                              if (acType.equalsIgnoreCase(acType4)) {
                                            	    System.out.println("Match found for Trip Number: " + tripNumber);

                                                  List<String> compatibleStands = findCompatibleStands(acType, "allAC.xlsx", "stand.xlsx", schengenStatus, excludedStands, airlineCode);
                                                  long diffHours = 0;
                                                  long diffMinutes = 0;
                                                  matchCount++; // Increment the match counter

                                                  try {
                                                      Date startTime = parseDateTime(scheduledDate, scheduledTime4);
                                                      Date endTime = parseDateTime(scheduledDate, scheduledTime);
                                                      long diffInMillis = endTime.getTime() - startTime.getTime();
                                                      diffHours = diffInMillis / (60 * 60 * 1000);
                                                      diffMinutes = (diffInMillis % (60 * 60 * 1000)) / (60 * 1000);
                                                  } catch (ParseException e) {
                                                      e.printStackTrace();
                                                      throw new RuntimeException("Error parsing date and time", e);
                                                  }
                                                  if (compatibleStands.isEmpty()) {
                                                      System.out.println("No available stands for " + acType + " in time slot " + scheduledTime + " - " + scheduledTime4);
                                                  
                                                  } else {
                                                      
                                                	  String reservedStand = getSequentialStand(compatibleStands);
                                                      System.out.println("Sequential Stand: " + reservedStand);
                                                      
                                                      boolean canReserve = true;
                                                      List<String> alternativeStands = findCompatibleStands(acType, "allAC.xlsx", "stand.xlsx", schengenStatus, excludedStands, airlineCode);

                                                   // Find another compatible stand without overlaps
                                                      String alternativeStand = null;
                                                      for (String stand : alternativeStands) {
                                                          boolean standIsFree = true;
                                                          for (FlightInfo existingReservation : standReservations.values()) {
                                                              if (isOverlapping(existingReservation, new FlightInfo(tripNumber, scheduledDate, scheduledDate2, scheduledTime, scheduledTime4, acType, acType4, aircraftCategory, "", stand, "YOUR_AIRPORT_CODE_HERE", routing, schengenStatus, airlineCode, scheduledDate3))) {
                                                                  standIsFree = false;
                                                                  break; // This stand also has an overlap
                                                              }
                                                          }
                                                          if (standIsFree) {
                                                              alternativeStand = stand; // Found a suitable stand without overlap
                                                              break;
                                                          }
                                                      }

                                                      // If no suitable stand is found (meaning all compatible stands have overlaps)
                                                      if (alternativeStand == null) {
                                                          System.out.println("Unable to resolve overlap for Trip " + tripNumber + ". No stand assigned.");
                                                      } else {
                                                          // Assigned new stand to resolve overlap
                                                          System.out.println("Assigned new stand " + alternativeStand + " to resolve overlap.");
                                                          // Reserve the alternative stand
                                                          reserveStand(tripNumber, scheduledDate, scheduledDate2, scheduledTime, scheduledTime4, acType, acType4, aircraftCategory, alternativeStand, "YOUR_AIRPORT_CODE_HERE", routing, schengenStatus, airlineCode, scheduledDate3);
                                                          System.out.println("Debug: Adding reservation for flight " + tripNumber + " with stand " + alternativeStand);
                                                          // Update reservations map
                                                          standReservations.put(tripNumber, new FlightInfo(tripNumber, scheduledDate, scheduledDate2, scheduledTime, scheduledTime4, acType, acType4, aircraftCategory, "", alternativeStand, "YOUR_AIRPORT_CODE_HERE", routing, schengenStatus, airlineCode, scheduledDate3));
                                                          System.out.println("Reserved Stand for " + acType + " (Trip " + tripNumber + "): " + alternativeStand +
                                                                  " | Time Slot: " + scheduledTime4 + " - " + scheduledTime);
                                                          System.out.println("Debug: Reserved stand " + alternativeStand + " for flight " + tripNumber + ". Current reservations count: " + standReservations.size());
                                                          System.out.println("Scheduled Date 3 (from matching row): " + scheduledDate3);
                                                    	  System.out.println("Scheduled Date (from base row): " + scheduledDate);                                                      
                                                          System.out.println("Scheduled Time : " + scheduledTime4);
                                                    	  System.out.println("Scheduled Time : " + scheduledTime);                                                      

                                                      }

                                                  }
                                              } else {
                                                  System.out.println("Aircraft types do not match for Trip " + tripNumber);
                                              }
                                              break;
                                          }
                                      }
                                  }
                              }
                          }
                      }
                      System.out.println("\nStand Reservations:");
                      List<Map.Entry<String, FlightInfo>> sortedReservations = new ArrayList<>(standReservations.entrySet());
                      System.out.println("Stand Reservations before sorting:");
                      for (Map.Entry<String, FlightInfo> entry : standReservations.entrySet()) {
                          FlightInfo reservationInfo = entry.getValue();
                          System.out.println("Trip " + entry.getKey() + ":");
                          System.out.println("Scheduled Date: " + reservationInfo.scheduledDate);
                          System.out.println("Scheduled Time: " + reservationInfo.scheduledTime4 + " - " + reservationInfo.scheduledTime);
                          System.out.println("Aircraft Type: " + reservationInfo.acType);
                          System.out.println("Aircraft Category: " + reservationInfo.aircraftCategory);
                          System.out.println("Reserved Stand: " + reservationInfo.reservedStand);
                          System.out.println();
                      }
                   // Sort the reservations by scheduledTime4
                      sortedReservations.sort((entry1, entry2) -> {
                    	    FlightInfo reservation1 = entry1.getValue();
                    	    FlightInfo reservation2 = entry2.getValue();
                    	    try {
                    	        // Combine date and time into one string and parse it into a Date object
                    	    	SimpleDateFormat sdf = new SimpleDateFormat("dd.MM.yyyy HH:mm");

                    	    	Date dateTime1 = sdf.parse(reservation1.scheduledDate3 + " " + reservation1.scheduledTime4);
                    	        Date dateTime2 = sdf.parse(reservation2.scheduledDate3 + " " + reservation2.scheduledTime4);

                    	        //  Date.compareTo() method to compare the two Date objects
                    	        return dateTime1.compareTo(dateTime2);
                    	    } catch (ParseException e) {
                    	        e.printStackTrace();
                    	        // Handle parsing error, perhaps by  comparison that ensures stability, like comparing by trip number
                    	        return reservation1.tripNumber.compareTo(reservation2.tripNumber);
                    	    }
                    	});

                    	// Print the sorted reservations as before
                    	System.out.println("Sorted Reservations:");
                    	for (Map.Entry<String, FlightInfo> entry : sortedReservations) {
                    	    FlightInfo reservationInfo = entry.getValue();
                    	    String tripNumber = entry.getKey();
                    	    System.out.println("Trip " + tripNumber + ":");
                    	    System.out.println("Reserved on: " + reservationInfo.scheduledDate3 + " " + reservationInfo.scheduledDate +
                    	    " " + reservationInfo.scheduledTime4 + " - " + reservationInfo.scheduledTime);
                    	    System.out.println("Aircraft Type: " + reservationInfo.acType);
                    	    System.out.println("Aircraft Category: " + reservationInfo.aircraftCategory);
                    	    System.out.println("Reserved Stand: " + reservationInfo.reservedStand);
                    	    System.out.println("Routing: " + reservationInfo.routing);
                    	    System.out.println("Airline Code: " + reservationInfo.airlineCode);
                    	    System.out.println("Schengen: " + reservationInfo.schengenStatus);
                    	}
                    

                      writeToExcel(new ArrayList<>(standReservations.entrySet()));

                      // Print the sorted reservations
                      System.out.print("Enter the flight number to cancel its reservation (or '0' to skip): ");
                      String flightToCancel = "";
                      try {
                          flightToCancel = scanner.nextLine().trim();
                      } catch (NoSuchElementException e) {
                          System.out.println("No input found for flight cancellation.");
                      }

                      if (!flightToCancel.equals("0") && !flightToCancel.isEmpty()) {
                    	  deleteReservation(flightToCancel);
                      }

                      // Reassign stands for flights
                      System.out.print("Enter the flight number to reassign its stand (or '0' to skip): ");
                      String flightToReassign = "";
                      try {
                          flightToReassign = scanner.nextLine().trim();
                      } catch (NoSuchElementException e) {
                          System.out.println("No input found for stand reassignment.");
                      }

                      if (!flightToReassign.equals("0") && !flightToReassign.isEmpty()) {
                          reassignStandForFlight(flightToReassign);
                      }


                      // Write updated reservations to Excel
                      writeToExcel(new ArrayList<>(standReservations.entrySet()));
                      // After populating standReservations map in main method
                      int numberOfReservations = standReservations.size();
                      //System.out.println("Number of flight reservations made: " + numberOfReservations);
                      //System.out.println("Number of matches found: " + matchCount);

                  } else {
                      System.out.println("No start date found for the desired date: " + desiredDate);
                  }
                  scanner.close();
              } else {
                  System.out.println("File not found!");
              }
          } catch (IOException | RuntimeException e) {
              e.printStackTrace();
          }
      }
      
      
      public static void deleteReservation(String flightNumber) {
    	    try {
    	        // Load the latest reservations before making changes
    	        loadReservationsFromExcel();

    	        FlightInfo flightInfo = standReservations.get(flightNumber);
    	        if (flightInfo != null) {
    	            System.out.println("Deleting reservation for Flight " + flightNumber + ": " + flightInfo.reservedStand);

    	            // Remove the reservation from the map
    	            standReservations.remove(flightNumber);

    	            // Update the Excel file to reflect the changes
    	            writeToExcel(new ArrayList<>(standReservations.entrySet()));
    	            System.out.println("Reservation deleted successfully for Flight " + flightNumber);
    	        } else {
    	            System.out.println("Flight " + flightNumber + " not found in reservations.");
    	        }
    	    } catch (IOException e) {
    	        System.err.println("Failed to load or write to Excel: " + e.getMessage());
    	    } catch (ParseException e) {
    	        System.err.println("Failed to parse dates while loading reservations: " + e.getMessage());
    	    }
    	}

      
   //  method to reassign stand for a specific flight
      public static void reassignStandForFlight(String flightNumber) {
    	    try {
    	        // Load the latest reservations before making changes
    	        loadReservationsFromExcel();
    	    } catch (IOException | ParseException e) {
    	        System.err.println("Failed to load reservations: " + e.getMessage());
    	        return; // Exit if unable to load data
    	    }

    	    FlightInfo flightInfo = standReservations.get(flightNumber);

    	    if (flightInfo != null) {
    	        System.out.println("Current Stand for Flight " + flightNumber + ": " + flightInfo.reservedStand);

    	        Set<String> excludedStands = new HashSet<>();
    	        excludedStands.add(flightInfo.reservedStand); // Exclude current stand from the options

    	        try {
    	            List<String> compatibleStands = findCompatibleStands(flightInfo.acType, "allAC.xlsx", "stand.xlsx", flightInfo.schengenStatus, excludedStands, flightInfo.airlineCode);

    	            if (compatibleStands.isEmpty()) {
    	                System.out.println("No available stands for reassignment.");
    	            } else {
    	                Scanner scanner = new Scanner(System.in);
    	                System.out.print("Enter the new stand number (or press Enter for automatic assignment): ");
    	                String newStand = scanner.nextLine().trim();

    	                if (newStand.isEmpty()) { // If user didn't provide a new stand, assign automatically
    	                    newStand = getSequentialStand(compatibleStands);
    	                }

    	                // Update the reservation with the new stand
    	                flightInfo.reservedStand = newStand;
    	                standReservations.put(flightNumber, flightInfo); // Update the map with the new information
    	                System.out.println("Stand reassigned successfully for Flight " + flightNumber + ": " + newStand);

    	                // Optionally, update the Excel file with the new reservations
    	                writeToExcel(new ArrayList<>(standReservations.entrySet()));
    	            }
    	        } catch (IOException e) {
    	            e.printStackTrace();
    	        }
    	    } else {
    	        System.out.println("Flight " + flightNumber + " not found in reservations.");
    	    }
    	}

    	private static void loadReservationsFromExcel() throws IOException, ParseException {
    	    InputStream inputStream = new FileInputStream("sorted_reservations.xlsx");
    	    Workbook workbook = new XSSFWorkbook(inputStream);
    	    Sheet sheet = workbook.getSheetAt(0);
    	    standReservations.clear(); // Clear existing data to avoid duplicates

    	    DataFormatter formatter = new DataFormatter();
    	    SimpleDateFormat dateFormat = new SimpleDateFormat("dd.MM.yyyy HH:mm");

    	    for (Row row : sheet) {
    	        if (row.getRowNum() == 0) continue; // Skip header row

    	        String tripNumber = formatter.formatCellValue(row.getCell(0));
    	        String airlineCode = tripNumber.substring(0, 2); // Assuming airline code is the first 2 chars
    	        tripNumber = tripNumber.substring(2);

    	        String dateTimeStr = formatter.formatCellValue(row.getCell(1));
    	        Date dateTime = dateFormat.parse(dateTimeStr);

    	        String dateTimeStr2 = formatter.formatCellValue(row.getCell(2));
    	        Date dateTime2 = dateFormat.parse(dateTimeStr2);

    	        String acType = formatter.formatCellValue(row.getCell(3));
    	        String aircraftCategory = formatter.formatCellValue(row.getCell(4));
    	        String reservedStand = formatter.formatCellValue(row.getCell(5)).replace("stand", ""); // Assuming stand prefix "stand"
    	        String routing = formatter.formatCellValue(row.getCell(6));
    	        String schengenStatus = formatter.formatCellValue(row.getCell(8));

    	        FlightInfo flightInfo = new FlightInfo(tripNumber, dateFormat.format(dateTime), dateFormat.format(dateTime2),
    	                "", "", acType, "", aircraftCategory, "", reservedStand, "YOUR_AIRPORT_CODE_HERE", routing, schengenStatus, airlineCode, dateFormat.format(dateTime));
    	        standReservations.put(tripNumber, flightInfo);
    	    }

    	    workbook.close();
    	    inputStream.close();
    	}

      
      

      
      private static Map<String, String> readFlightCodes(String filename) throws IOException {
          Map<String, String> flightCodes = new HashMap<>();
          // logic to read flight codes from the specified file (ACall.xlsx)
          return flightCodes;
      }
     
      private static void writeToExcel(List<Map.Entry<String, FlightInfo>> sortedReservations) {
   	    try {
   	        Workbook workbook = new XSSFWorkbook();
   	        Sheet sheet = workbook.createSheet("Sorted Reservations");
   	        // Create headers
   	        Row headerRow = sheet.createRow(0);
   	        String[] headers = {"Trip Number", "Scheduled Date", "Scheduled Date 2", "Aircraft Type", "Aircraft Category", "Reserved Stand", "Routing", "Airline", "Schengen"};
   	        for (int i = 0; i < headers.length; i++) {
   	            Cell cell = headerRow.createCell(i);
   	            cell.setCellValue(headers[i]);
   	        }
   	        // Write data
   	        int rowNum = 1;
   	        for (Map.Entry<String, FlightInfo> entry : sortedReservations) {
   	            FlightInfo reservationInfo = entry.getValue();
   	            String tripNumber = entry.getKey();
   	            Row row = sheet.createRow(rowNum++);
   	            row.createCell(0).setCellValue(reservationInfo.airlineCode + tripNumber);
   	            
   	         String dateTime = reservationInfo.scheduledDate3 + " " + reservationInfo.scheduledTime4;
             Cell dateCell = row.createCell(1);
             dateCell.setCellValue(dateTime);
             String dateTime2 = reservationInfo.scheduledDate + " " + reservationInfo.scheduledTime;
             Cell dateCell2 = row.createCell(2);
             dateCell2.setCellValue(dateTime2);
             
             CellStyle cellStyle = workbook.createCellStyle();
             CreationHelper createHelper = workbook.getCreationHelper();
             cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd.MM.yyyy HH:mm"));
             dateCell.setCellStyle(cellStyle);
             
   	            //row.createCell(1).setCellValue(reservationInfo.scheduledDate3);
   	            //row.createCell(2).setCellValue(reservationInfo.scheduledDate);
   	            //row.createCell(3).setCellValue(reservationInfo.scheduledTime4);
   	            //row.createCell(4).setCellValue(reservationInfo.scheduledTime);
   	            
   	            row.createCell(3).setCellValue(reservationInfo.acType);
   	            row.createCell(4).setCellValue(reservationInfo.aircraftCategory);
   	            row.createCell(5).setCellValue("stand"+reservationInfo.reservedStand);
   	            row.createCell(6).setCellValue(reservationInfo.routing);
   	            row.createCell(7).setCellValue(reservationInfo.airlineCode);
   	            row.createCell(8).setCellValue(reservationInfo.schengenStatus);
   	        }
   	        // Write the workbook content to a file
   	        FileOutputStream fileOut = new FileOutputStream("sorted_reservations.xlsx");
   	        workbook.write(fileOut);
   	        fileOut.close();
   	        workbook.close();
   	        System.out.println("Sorted reservations written to sorted_reservations.xlsx");
   	    } catch (IOException e) {
   	        e.printStackTrace();
   	    }
   	}
   
      
      
      private static Map<String, List<String>> readAirlinePreferences(String filename) throws IOException {
          Map<String, List<String>> airlinePreferences = new HashMap<>();
          InputStream inputStream = Save7.class.getClassLoader().getResourceAsStream(filename);
          if (inputStream == null) {
              throw new IOException("File not found: " + filename);
          }
          Workbook workbook = new XSSFWorkbook(inputStream);
          Sheet sheet = workbook.getSheetAt(0);
          
          for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
              Row row = sheet.getRow(rowIndex);
              Cell airlineCell = row.getCell(0); // Airline name in the first column
              if (airlineCell == null) continue; // Skip rows where the first cell (airline name) is empty
              String airlineName = getCellStringValue(airlineCell);
              List<String> preferences = new ArrayList<>();
              for (int i = 1; i <= 16; i++) { // Loop through columns 2 to 9 for stand preferences
                  Cell preferenceCell = row.getCell(i);
                  if (preferenceCell != null) {
                      String stand = getCellStringValue(preferenceCell);
                      if (!stand.isEmpty()) {
                          preferences.add(stand);
                      }
                  }
              }
              airlinePreferences.put(airlineName, preferences);
              // Directly print each airline's preferences
              //System.out.print("Airline: " + airlineName + " - Preferences: ");
              for (int i = 0; i < preferences.size(); i++) {
                  //System.out.print(preferences.get(i) + (i < preferences.size() - 1 ? ", " : ""));
              }
              //System.out.println();
          }
          workbook.close();
          inputStream.close();
          return airlinePreferences;
  }
      
     
      private static String getCellStringValue(Cell cell) {
    	    switch (cell.getCellType()) {
    	        case STRING:
    	            return cell.getStringCellValue().trim();
    	        case NUMERIC:
    	            // Check if the numeric value is an integer
    	            double numericValue = cell.getNumericCellValue();
    	            if (Math.floor(numericValue) == numericValue) {
    	                // It's an integer, format without the decimal part
    	                return String.format("%d", (int)numericValue);
    	            } else {
    	                // It's a real decimal number, keep it as is
    	                return String.format("%s", numericValue);
    	            }
    	        default:
    	            return "";
    	    }
    	}

      
private static Map<String, String> readACallData(String filename) throws IOException {
  Map<String, String> airportData = new HashMap<>();
  InputStream inputStream = Save7.class.getClassLoader().getResourceAsStream(filename);
  if (inputStream != null) {
      Workbook workbook = new XSSFWorkbook(inputStream);
      Sheet sheet = workbook.getSheetAt(0);
      for (Row row : sheet) {
          Cell cellRouting = row.getCell(0); // Column A (Routing code)
          Cell cellSchengen = row.getCell(4); // Column E (Schengen status)
          if (cellRouting != null && cellSchengen != null) {
              String routingCode = cellRouting.getStringCellValue().trim();
              String schengenStatus = cellSchengen.getStringCellValue().trim();
              airportData.put(routingCode, schengenStatus);
          }
      }
  }
  return airportData;
}


   // Method to read airport codes from letovyradcas.xlsx
      private static List<String> readLetovyradcasData(String filename) throws IOException {
          List<String> airportCodes = new ArrayList<>();
          InputStream inputStream = Save7.class.getClassLoader().getResourceAsStream(filename);
          if (inputStream != null) {
              Workbook workbook = new XSSFWorkbook(inputStream);
              Sheet sheet = workbook.getSheetAt(0);
              for (Row row : sheet) {
                  Cell cellAirportCode = row.getCell(12); // Assuming airport codes are in column M (index 12)
                  if (cellAirportCode != null) {
                      String airportCode = cellAirportCode.getStringCellValue().trim();
                      airportCodes.add(airportCode);
                  }
              }
          }
          return airportCodes;
      }
  public static Date parseDateTime(String date, String time) throws ParseException {
      SimpleDateFormat dateFormat = new SimpleDateFormat("dd.MM.yyyy HH:mm");
      String dateTimeString = date + " " + time;
      return dateFormat.parse(dateTimeString);
  }
private static String getACType(Cell cell) {
   if (cell != null) {
       cell.setCellType(CellType.STRING);
       return cell.getStringCellValue().trim();
   }
   return "";
}
private static String findStartDate(Sheet sheet, String desiredDate) {
   for (int i = 1; i <= sheet.getLastRowNum(); i++) {
       Row row = sheet.getRow(i);
       Cell cellScheduledDate = row.getCell(0);
       String scheduledDate = getStringValue(cellScheduledDate);
       if (scheduledDate.equals(desiredDate)) {
           // If the scheduled date matches the desired date, return the start date
           Cell cellStartDate = row.getCell(6);
           return getStringValue(cellStartDate);
       }
   }
   return null;
}
private static String getCategoryFromAllAC(String acType, String allACFilename) throws IOException {
   InputStream inputStream = zk.class.getClassLoader().getResourceAsStream(allACFilename);
   if (inputStream != null) {
       Workbook workbook = new XSSFWorkbook(inputStream);
       Sheet sheet = workbook.getSheetAt(0);
       for (Row row : sheet) {
           Cell cellACType = row.getCell(0);
           Cell cellCategory = row.getCell(1);
           if (acType.equalsIgnoreCase(getStringValue(cellACType))) {
               return getStringValue(cellCategory);
           }
       }
   }
   return null;
}


private static List<String> findCompatibleStands(String acType, String allACFilename, String standFilename, String schengenStatus, Set<String> excludedStands, String airlineCode) throws IOException {
    Map<String, List<String>> airlinePreferences = readAirlinePreferences("preferences.xlsx");
    InputStream inputStream = Save7.class.getClassLoader().getResourceAsStream(allACFilename);
    Workbook allACWorkbook = new XSSFWorkbook(inputStream);
    InputStream standInputStream = Save7.class.getClassLoader().getResourceAsStream(standFilename);
    Workbook standWorkbook = new XSSFWorkbook(standInputStream);
    List<String> compatibleStands = new ArrayList<>();
    Map<String, Double> standWingSpanMap = new HashMap<>(); // Map to store stand and its wingspan
    
    if (inputStream != null && standInputStream != null) {
        Sheet allACSheet = allACWorkbook.getSheetAt(0);
        Sheet standSheet = standWorkbook.getSheetAt(0);
        double acWingSpan = getAircraftWingSpan(acType, allACSheet);
        for (Row row : standSheet) {
            Cell cellStand = row.getCell(0);
            String standId = getStringValue(cellStand);
            if (excludedStands.contains(standId)) {
                continue;
            }
            Cell cellWingSpan = row.getCell(1);
            Cell cellSchengenStatus = row.getCell(3);
            double standWingSpan = getNumericValue(cellWingSpan);
            String standSchengenStatus = getStringValue(cellSchengenStatus);
            if (acWingSpan <= standWingSpan && standSchengenStatus.equals(schengenStatus)) {
                compatibleStands.add(standId);
                standWingSpanMap.put(standId, standWingSpan); // Store wingspan for later sorting
            }
        }
    }

    List<String> sortedStands = new ArrayList<>(); // This will hold the final sorted list of stands
    // Prioritize stands based on airline preferences
    if (airlinePreferences.containsKey(airlineCode)) {
        List<String> preferredStands = airlinePreferences.get(airlineCode);

        // Add preferred and compatible stands first
        for (String preferredStand : preferredStands) {
            if (compatibleStands.contains(preferredStand)) {
                sortedStands.add(preferredStand);
                compatibleStands.remove(preferredStand); // Remove the stand from compatibleStands to avoid duplication
            }
        }
    }

    // Sort the rest of the compatible stands by wingspan
    compatibleStands.sort(Comparator.comparing(stand -> standWingSpanMap.get(stand)));

    // Add the sorted compatible stands to the sortedStands list
    sortedStands.addAll(compatibleStands);

    return sortedStands;
}



private static List<String> getSmallestStands(List<String> stands, Sheet standSheet) {
  List<String> smallestStands = new ArrayList<>();
  for (String stand : stands) {
      for (Row row : standSheet) {
          Cell cellStand = row.getCell(0);
          if (getStringValue(cellStand).equals(stand)) {
              smallestStands.add(stand);
              break;
          }
      }
  }
  return smallestStands;
}
private static double getAircraftWingSpan(String acType, Sheet allACSheet) {
   for (Row row : allACSheet) {
       Cell cellACType = row.getCell(0);
       Cell cellWingSpan = row.getCell(1);
       String sheetACType = getStringValue(cellACType);
       double sheetWingSpan = getNumericValue(cellWingSpan);
       if (acType.equalsIgnoreCase(sheetACType)) {
           //System.out.println("Match found! Wing Span: " + sheetWingSpan);
           return sheetWingSpan;
       }
   }
   return 0.0;
}
private static void reserveStand(String tripNumber, String scheduledDate, String scheduledDate2, String scheduledTime,
                                 String scheduledTime4, String acType, String acType4, String aircraftCategory,
                                 String reservedStand, String airportCode, String routing, String schengenStatus, String airlineCode, String scheduledDate3) {
    List<String> compatibleStands;
    try {
        //return a list of stand IDs that are compatible with the aircraft type and Schengen status.
    	Set<String> excludedStands = new HashSet<>(); // Initialize an empty set if there re no stands to exclude initially

    	compatibleStands = findCompatibleStands(acType, "allAC.xlsx", "stand.xlsx", schengenStatus, excludedStands, airlineCode);
    } catch (IOException e) {
        System.err.println("Failed to find compatible stands due to an IO error: " + e.getMessage());
        return;
    }

    if (compatibleStands.isEmpty()) {
        System.out.println("No available stands for " + acType + " in time slot " + scheduledTime4 + " - " + scheduledTime);
        return;
    }

    
    String finalStand = reservedStand;
    FlightInfo tempReservation = new FlightInfo(tripNumber, scheduledDate, scheduledDate2, scheduledTime,
            scheduledTime4, acType, acType4, aircraftCategory, "", reservedStand, airportCode, routing, schengenStatus, airlineCode, scheduledDate3);

    // Initial check for overlaps with the desired stand
    for (FlightInfo existingReservation : standReservations.values()) {
        if (existingReservation.reservedStand.equals(finalStand) && isOverlapping(existingReservation, tempReservation)) {
            System.out.println("Overlap detected for stand " + finalStand + " with Trip " + existingReservation.tripNumber);
            finalStand = null; // Clear the finalStand as a signal to find a new one
            break;
        }
    }

    // If an overlap was found or the original stand was unavailable, try to find a new stand
    if (finalStand == null) {
        for (String stand : compatibleStands) {
            boolean standIsFree = true;
            for (FlightInfo existingReservation : standReservations.values()) {
                if (existingReservation.reservedStand.equals(stand) && isOverlapping(existingReservation, tempReservation)) {
                    standIsFree = false;
                    break;
                }
            }
            if (standIsFree) {
                finalStand = stand; // Found a suitable stand without overlap
                System.out.println("Assigned new stand " + finalStand + " to Trip " + tripNumber + " to resolve overlap.");
                break;
            }
        }
    }

    // If no suitable stand is found (meaning all compatible stands have overlaps)
    if (finalStand == null) {
        System.out.println("Unable to resolve overlap for Trip " + tripNumber + ". No stand assigned.");
        return; // Exiting the method as no stand could be assigned
    }

    // Update the temporary reservation with the final stand, whether it was the originally requested or a new one
    tempReservation.reservedStand = finalStand;

    // Check if the reservation for this trip number already exists and remove it before adding a new one
    if (standReservations.containsKey(tripNumber)) {
        standReservations.remove(tripNumber);
    }

    // Update the global reservations map
    standReservations.put(tripNumber, tempReservation);

    // Log the final assignment and increment reservation count only if a new stand is assigned
    System.out.println("Reservation confirmed for Trip " + tripNumber + ": Stand " + finalStand + ", Time Slot: " + scheduledTime4 + " - " + scheduledTime);
    if (!finalStand.equals(reservedStand)) {
        System.out.println("Debug: Reserved stand " + finalStand + " for flight " + tripNumber + ". Current reservations count: " + standReservations.size());
    }
}




// Method to check if the flight is Schengen or Non-Schengen based on ACall.xlsx
private static boolean isSchengenFlight(String acType, String aCallFilename) throws IOException {
  InputStream inputStream = Save7.class.getClassLoader().getResourceAsStream(aCallFilename);
  if (inputStream != null) {
      Workbook workbook = new XSSFWorkbook(inputStream);
      Sheet sheet = workbook.getSheet("SCH");
      for (Row row : sheet) {
          Cell cellACType = row.getCell(0);
          Cell cellSchengen = row.getCell(4);
          if (acType.equalsIgnoreCase(getStringValue(cellACType))) {
              String schengenValue = getStringValue(cellSchengen);
              return "SCHENGEN".equalsIgnoreCase(schengenValue);
          }
      }
  }
  return false;
}

// Method to check if the destination is Schengen or Non-Schengen based on letovyradcas.xlsx
private static boolean isSchengenDestination(String tripNumber, String letovyradcasFilename) throws IOException {
  InputStream inputStream = Save7.class.getClassLoader().getResourceAsStream(letovyradcasFilename);
  if (inputStream != null) {
      Workbook workbook = new XSSFWorkbook(inputStream);
      Sheet sheet = workbook.getSheetAt(0);
      for (Row row : sheet) {
          Cell cellTripNumber = row.getCell(8);
          Cell cellSchengen = row.getCell(6);
          if (tripNumber.equalsIgnoreCase(getStringValue(cellTripNumber))) {
              String schengenValue = getStringValue(cellSchengen);
              return "SCHENGEN".equalsIgnoreCase(schengenValue);
          }
     
      }
  }
  return false;
}

private static Map<String, List<String>> readStandRestrictions(String filename) throws IOException {
    System.out.println("Reading stand restrictions from: " + filename);
    Map<String, List<String>> standRestrictions = new HashMap<>();
    InputStream inputStream = Save7.class.getClassLoader().getResourceAsStream(filename);
    if (inputStream == null) {
        System.err.println("Error: File not found - " + filename);
        throw new IOException("File not found: " + filename);
    }
    Workbook workbook = new XSSFWorkbook(inputStream);
    Sheet sheet = workbook.getSheetAt(0);

    DataFormatter formatter = new DataFormatter(); // Create a formatter to convert numeric cells to strings

    int rowNumber = 0;
    for (Row row : sheet) {
        rowNumber++;
        Cell cellStand = row.getCell(0); // Stand name in column A
        if (cellStand == null) {
            System.out.println("Skipping row " + rowNumber + ": Stand name cell is empty.");
            continue;
        }
        String standName = formatter.formatCellValue(cellStand).trim(); // Use formatter here
        System.out.println("Processing stand: " + standName + " (Row " + rowNumber + ")");

        List<String> restrictedStands = new ArrayList<>();
        for (int i = 1; i <= 4; i++) { // Loop from column B to E
            Cell cell = row.getCell(i);
            if (cell != null) {
                String restrictedStand = formatter.formatCellValue(cell).trim(); // Use formatter here
                if (!restrictedStand.isEmpty()) {
                    restrictedStands.add(restrictedStand);
                    System.out.println(" - Found restriction with stand: " + restrictedStand);
                }
            }
        }

        if (!restrictedStands.isEmpty()) {
            System.out.println(" - Total restrictions found for " + standName + ": " + restrictedStands.size());
        } else {
            System.out.println(" - No restrictions found for " + standName);
        }

        standRestrictions.put(standName, restrictedStands);
    }

    workbook.close();
    inputStream.close();
    System.out.println("Completed reading stand restrictions. Total stands processed: " + standRestrictions.size());
    return standRestrictions;
}



private static boolean isOverlapping(FlightInfo flightInfo1, FlightInfo flightInfo2) {
    SimpleDateFormat sdf = new SimpleDateFormat("dd.MM.yyyy HH:mm");
    try {
        // Parse the arrival and departure for both flights
        Date arrival1 = sdf.parse(flightInfo1.scheduledDate3 + " " + flightInfo1.scheduledTime4);
        Date departure1 = sdf.parse(flightInfo1.scheduledDate + " " + flightInfo1.scheduledTime);
        Date arrival2 = sdf.parse(flightInfo2.scheduledDate3 + " " + flightInfo2.scheduledTime4);
        Date departure2 = sdf.parse(flightInfo2.scheduledDate + " " + flightInfo2.scheduledTime);

        // Adjust for overnight conditions
        if (departure1.before(arrival1)) {
            Calendar c = Calendar.getInstance();
            c.setTime(departure1);
            c.add(Calendar.DATE, 1);
            departure1 = c.getTime();
        }
        if (departure2.before(arrival2)) {
            Calendar c = Calendar.getInstance();
            c.setTime(departure2);
            c.add(Calendar.DATE, 1);
            departure2 = c.getTime();
        }

        // Check direct stand overlap
        if (flightInfo1.reservedStand.equals(flightInfo2.reservedStand)) {
            if ((arrival1.before(departure2) && departure1.after(arrival2)) || (arrival2.before(departure1) && departure2.after(arrival1))) {
                System.out.println("Direct overlap detected on stand " + flightInfo1.reservedStand);
                return true;
            }
        }

        // Check restricted stands overlap
        List<String> restrictionsForStand1 = standRestrictions.getOrDefault(flightInfo1.reservedStand, Collections.emptyList());
        List<String> restrictionsForStand2 = standRestrictions.getOrDefault(flightInfo2.reservedStand, Collections.emptyList());
        
        if (restrictionsForStand1.contains(flightInfo2.reservedStand) || restrictionsForStand2.contains(flightInfo1.reservedStand)) {
            if ((arrival1.before(departure2) && departure1.after(arrival2)) || (arrival2.before(departure1) && departure2.after(arrival1))) {
                System.out.println("Overlap detected due to restricted stands between " + flightInfo1.reservedStand + " and " + flightInfo2.reservedStand);
                return true; // Overlap detected due to stand restrictions
            }
        }

    } catch (ParseException e) {
        System.err.println("Error parsing date/time: " + e.getMessage());
        return false;
    }
    System.out.println("No overlap detected for stand " + flightInfo1.reservedStand + " or related restricted stands.");
    return false; // No overlap detected
}



private static String cellToString(Cell cell) {
    if (cell == null) return "";

    switch (cell.getCellType()) {
        case STRING:
            return cell.getStringCellValue();
        case NUMERIC:
            // For numeric, convert to String
            return new DataFormatter().formatCellValue(cell);
        case BOOLEAN:
            // Convert boolean to String
            return Boolean.toString(cell.getBooleanCellValue());
        case FORMULA:
            return cell.getCellFormula();
        default:
            return "";
    }
}




private static double getNumericValue(Cell cell) {
   if (cell != null && cell.getCellType() == CellType.NUMERIC) {
       return cell.getNumericCellValue();
   }
   return 0.0;
}
private static String getStringValue(Cell cell) {
   if (cell != null) {
       cell.setCellType(CellType.STRING);
       return cell.getStringCellValue().trim();
   }
   return "";
}

private static int currentIndex = 0; // Keep track of the current index for sequential selection

private static String getSequentialStand(List<String> stands) {
    if (stands != null && !stands.isEmpty()) {
        if (currentIndex >= stands.size()) {
            currentIndex = 0; // Reset the index if it exceeds the list size
        }
        String stand = stands.get(currentIndex);
        currentIndex++; // Move to the next index for the next call
        return stand;
    } else {
        return "NoCompatibleStand";
    }
}


private static String getSchengenStatus(String flightCode, Map<String, String> airportData) {
  //System.out.println("Flight Code: " + flightCode);
   // Check if the flight code exists in the airport data
  if (airportData.containsKey(flightCode)) {
      String schengenStatus = airportData.get(flightCode);
      //System.out.println("Schengen Status Retrieved: " + schengenStatus);
    
      return "SCHENGEN".equalsIgnoreCase(schengenStatus) ? "SCHENGEN" : "NONSCHENGEN";
  } else {
      // If flight code information is not found, default to NONSCHENGEN
      return "NONSCHENGEN";
  	}
	}
}


	
}
