/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.program;

import com.excel.lib.util.Xls_Reader;
import java.awt.AWTException;
import java.awt.HeadlessException;
import java.awt.Robot;
import java.awt.event.InputEvent;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Scanner;
import lc.kra.system.keyboard.GlobalKeyboardHook;
import lc.kra.system.keyboard.event.GlobalKeyAdapter;
import lc.kra.system.keyboard.event.GlobalKeyEvent;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author DELL
 */
public class Test {
       private static boolean run = true;

    ArrayList<String> actions = new ArrayList<>();
    ArrayList<String> robotactions = new ArrayList<>();
    ArrayList<String> results = new ArrayList<>();
    static ArrayList<String> nativeactions;

    SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/YYYY");
    Date date = new Date();
    String todaysdate = formatter.format(date);

    String filename = "";
    String filepath = "";

    public static void main(String[] args) throws IOException, InterruptedException, AWTException {
        GlobalKeyboardHook keyboardHook = new GlobalKeyboardHook(true);

        System.out.println("Global keyboard hook successfully started, press [escape] key to shutdown. Connected keyboards:");

        keyboardHook.addKeyListener(new GlobalKeyAdapter() {

            @Override
            public void keyPressed(GlobalKeyEvent event) {
                String keychar = event.getKeyChar() + "";
                System.out.println(keychar.toUpperCase() + " Key Pressed");
                add_to_nativeactions(keychar.toUpperCase() + " Key Pressed");

                if (event.getVirtualKeyCode() == GlobalKeyEvent.VK_ESCAPE) {
                    run = false;
                }
            }

            @Override
            public void keyReleased(GlobalKeyEvent event) {
                String keychar = event.getKeyChar() + "";
                System.out.println(event.getKeyChar() + " Key Released");
                add_to_nativeactions(keychar.toUpperCase() + " Key Released");
            }
        });

        Test ee = new Test();

        ee.create_new_excel_file();
        Thread.sleep(500);
        ProcessBuilder pb = new ProcessBuilder("C:\\Program Files (x86)\\Attendance Management System\\AttendanceTracker11.8\\AttendanceTracker11.8.exe");
        runProcess(pb);

        int num = ee.get_last_empty_row();
        System.out.println("Last Non Empty Row:" + num);
        Thread.sleep(1500);
        ee.write_date(num + 2);
        ee.write_actions_excel(num + 2);
        try {
            ee.process_robot();

        } catch (Exception e) {
            e.printStackTrace();
        } finally {

            Thread.sleep(1500);
            ee.check_write_status(num + 2);
            ee.setfillcolor(num + 1);

        }

        System.exit(0);

    }

    public static void add_to_nativeactions(String text) {

        nativeactions.add(text);
    }

    public Test() throws HeadlessException {
        nativeactions = new ArrayList<>();
        nativeactions.add("Attendance Manager Opened");

    }

    public synchronized void check_write_status(int rownum) throws IOException {

        Xls_Reader reader = new Xls_Reader(filepath);
        int index = 0;

        for (int i = 0; i < actions.size(); i++) {
            if (actions.get(i).equalsIgnoreCase(robotactions.get(i))) {
                results.add("Yes");
            } else {
                results.add("No");
            }
        }

        System.out.println("Verification Results ArrayList:" + results);

        /*if (i < robotactions.size() && actions.get(i).equalsIgnoreCase(robotactions.get(i))) {
                results.add("yes");
            } else {
                results.add("no");
            }*/
        for (int i = 0; i < results.size(); i++) {
            reader.setCellData("Sheet1", "Status", rownum + i, results.get(i));

        }

    }

    public synchronized void setfillcolor(int rownum) throws FileNotFoundException, IOException {

        FileInputStream fis = new FileInputStream(filepath);
        FileOutputStream fos = null;

        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheet("Sheet1");

        CellStyle greencellStyle = workbook.createCellStyle();
        greencellStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        greencellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        CellStyle redcellStyle = workbook.createCellStyle();
        redcellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
        redcellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        for (int i = 0; i < results.size(); i++) {

            Row row = sheet.getRow(rownum + i);
            try {
                Cell cell1 = row.getCell(2);

                if (results.get(i).equalsIgnoreCase("yes")) {
                    cell1.setCellValue(results.get(i));
                    cell1.setCellStyle(greencellStyle);

                } else {
                    cell1.setCellValue(results.get(i));
                    cell1.setCellStyle(redcellStyle);
                }

            } catch (Exception e) {
                e.printStackTrace();
            }
            //cell1.setCellValue(results.get(i));
            //cell1.setCellStyle(cellStyle);

            fos = new FileOutputStream(filepath);
            workbook.write(fos);

        }
        fis.close();
        fos.close();
    }

    public void create_new_excel_file() {
        formatter = new SimpleDateFormat("dd-MM-YYYY");
        date = new Date();
        todaysdate = formatter.format(date);

        filename = "robotlogs_" + todaysdate + ".xlsx";

        
        try {
            XSSFWorkbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Sheet1");
            filepath = "E:\\javarobotfile\\" + filename + "";
            FileOutputStream fos = new FileOutputStream(filepath);
            workbook.write(fos);
            System.out.println("Excel File " + filename + " created successfully");
        } catch (FileNotFoundException ex) {
            ex.printStackTrace();
        }
        catch (IOException ex) {
            ex.printStackTrace();
        }

        Xls_Reader reader = new Xls_Reader(filepath);

        reader.addColumn("Sheet1", "Date");
        reader.addColumn("Sheet1", "Action Performed");
        reader.addColumn("Sheet1", "Status");

    }

    public synchronized int get_last_empty_row() throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook(filepath);
        Sheet sheet = workbook.getSheet("Sheet1");
        int row = sheet.getLastRowNum();
        workbook.close();
        return row;

    }

    public synchronized void write_date(int rownum) {

        Xls_Reader reader = new Xls_Reader(filepath);

        int rowcount = reader.getRowCount("Sheet1");
        SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/YYYY");
        
        Date date = new Date();
        String todaysdate = formatter.format(date);

        reader.setCellData("Sheet1", "Date", rownum, todaysdate);
        System.out.println("Date Entered at Row:" + rowcount);
    }

    private static void runProcess(ProcessBuilder pb) throws IOException {
        pb.redirectErrorStream(true);
        pb.directory(new File("C:\\Program Files (x86)\\Attendance Management System\\AttendanceTracker11.8"));
        Process p = pb.start();
        System.out.println("Attendance Manager Opened");

        /*BufferedReader reader = new BufferedReader(new InputStreamReader(p.getInputStream()));
    String line;
    while ((line = reader.readLine()) != null) {
        System.out.println(line);*/
    }

    private static boolean isProcessRunning(String processName) throws IOException, InterruptedException {
        ProcessBuilder processBuilder = new ProcessBuilder("tasklist.exe");
        Process process = processBuilder.start();
        String tasksList = toString(process.getInputStream());
        return tasksList.contains(processName);
    }

    private static String toString(InputStream inputStream) {
        Scanner scanner = new Scanner(inputStream, "UTF-8").useDelimiter("\\A");
        String string = scanner.hasNext() ? scanner.next() : "";
        scanner.close();

        return string;
    }

    public boolean Validation_AttendanceM_Opened() throws AWTException, IOException, InterruptedException {
        boolean opened = false;
        File application = new File("C:\\Program Files (x86)\\Attendance Management System\\AttendanceTracker11.8\\AttendanceTracker11.8.exe");
        String applicationName = application.getName();

        if (!isProcessRunning(applicationName)) {
            //Desktop.getDesktop().open(application);
            System.out.println("Attendance Manager Not Opened");

        } else {
            System.out.println("Attendance Manager  Opened");
            opened = true;
        }

        return opened;

    }

    public synchronized void write_actions_excel(int rownum) {
//total 25 actions
        Xls_Reader reader = new Xls_Reader(filepath);

        actions.add("Attendance Manager Opened");
        actions.add("focusing on username textfield");
        actions.add("Entering Username");
        actions.add("Username Entered Successfully");
        actions.add("focusing on password textfield");
        actions.add("Entering password");
        actions.add("password Entered Successfully");
        actions.add("pressing login button");
        actions.add("opening data transfer window");
        actions.add("setting data transfer mode to usb connect");
        actions.add("pressing connect to connect to machine");
        actions.add("pressing download data button");
        actions.add("pressing yes on process attendance window");
        actions.add("pressing ok on data process complete");
        actions.add("pressing close button");
        actions.add("clicking on data processing");
        actions.add("clicking on process button");
        actions.add("clicking on ok button after processing");
        actions.add("clicking close button after data processing");
        actions.add("clicking on administration tab on menu bar");
        actions.add("clicking on database backup from administarion menu");
        actions.add("clicking on update database button");
        actions.add("clicking ok button after updating database");
        actions.add("clicking close button after updating database");
        actions.add("closing application");
        actions.add("Process Completed");

        for (int i = 0; i < actions.size(); i++) {
            reader.setCellData("Sheet1", "Action Performed", rownum + i, actions.get(i));
        }

    }

    public void process_robot() throws InterruptedException, AWTException, IOException {
        Robot robot = new Robot();

        Thread.sleep(8000);

        if (Validation_AttendanceM_Opened() == true) {
            robotactions.add("Attendance Manager Opened");
        } else {
            robotactions.add("Action Failed");
        }

        if (Validation_AttendanceM_Opened() == true) {
            robot.mouseMove(699, 362);//focusing on username textfield
            robotactions.add("focusing on username textfield");
            Thread.sleep(1000);
            robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
            Thread.sleep(200);
            robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
            Thread.sleep(200);

        } else {
            robotactions.add("Action Failed");
        }

        if (Validation_AttendanceM_Opened() == true) {
            robot.keyPress(KeyEvent.VK_A);
            robotactions.add("Entering Username");

            Thread.sleep(1000);

            robot.keyRelease(KeyEvent.VK_A);

            Thread.sleep(1000);

            robot.keyPress(KeyEvent.VK_D);

            Thread.sleep(1000);

            robot.keyRelease(KeyEvent.VK_D);

            Thread.sleep(1000);

            robot.keyPress(KeyEvent.VK_M);

            Thread.sleep(1000);

            robot.keyRelease(KeyEvent.VK_M);

            Thread.sleep(1000);

            robot.keyPress(KeyEvent.VK_I);

            Thread.sleep(1000);

            robot.keyRelease(KeyEvent.VK_I);

            Thread.sleep(1000);

            robot.keyPress(KeyEvent.VK_N);
            Thread.sleep(1000);

            robot.keyRelease(KeyEvent.VK_N);

            Thread.sleep(1000);

        } else {
            robotactions.add("Action Failed");
        }

        if (Validation_AttendanceM_Opened() == true) {
            robotactions.add("Username Entered Successfully");
        } else {
            robotactions.add("Action Failed");
        }


        /* robot.keyPress(KeyEvent.VK_TAB);
   System.out.println("Tab Key Pressed");
   Thread.sleep(1000);
   
   robot.keyRelease(KeyEvent.VK_T);
   System.out.println("Tab Key Released");
   Thread.sleep(1000);
         */
        if (Validation_AttendanceM_Opened() == true) {
            robot.mouseMove(752, 398);//focusing on password textfield

            robotactions.add("focusing on password textfield");
            Thread.sleep(1000);
            robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
            Thread.sleep(200);
            robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
            Thread.sleep(200);

        } else {
            robotactions.add("Action Failed");
        }

        if (Validation_AttendanceM_Opened() == true) {
            robot.keyPress(KeyEvent.VK_A);
            robotactions.add("Entering password");
            Thread.sleep(1000);

            robot.keyRelease(KeyEvent.VK_A);
            Thread.sleep(1000);

            robot.keyPress(KeyEvent.VK_D);
            Thread.sleep(1000);

            robot.keyRelease(KeyEvent.VK_D);
            Thread.sleep(1000);

            robot.keyPress(KeyEvent.VK_M);
            Thread.sleep(1000);

            robot.keyRelease(KeyEvent.VK_M);
            Thread.sleep(1000);

            robot.keyPress(KeyEvent.VK_I);
            Thread.sleep(1000);

            robot.keyRelease(KeyEvent.VK_I);
            Thread.sleep(1000);

            robot.keyPress(KeyEvent.VK_N);
            Thread.sleep(1000);

            robot.keyRelease(KeyEvent.VK_N);
            Thread.sleep(1000);

        } else {
            robotactions.add("Action Failed");
        }

        if (Validation_AttendanceM_Opened() == true) {
            robotactions.add("password Entered Successfully");
        } else {
            robotactions.add("Action Failed");
        }

        /*robot.keyPress(KeyEvent.VK_TAB);
   System.out.println("Tab Key Pressed");
   Thread.sleep(1000);
   
   robot.keyRelease(KeyEvent.VK_T);
   System.out.println("Tab Key Released");
   Thread.sleep(1000);
   
   robot.keyPress(KeyEvent.VK_ENTER);
   System.out.println("Enter Key Pressed");
   Thread.sleep(1000);
   
   robot.keyRelease(KeyEvent.VK_ENTER);
   System.out.println("Enter Key Released");
   Thread.sleep(1000);*/
        if (Validation_AttendanceM_Opened() == true) {
            robot.mouseMove(696, 439);//pressing login button
            robotactions.add("pressing login button");
            Thread.sleep(1000);
            robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
            Thread.sleep(200);
            robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
            Thread.sleep(5000);

        } else {
            robotactions.add("Action Failed");
        }

        if (Validation_AttendanceM_Opened() == true) {
            robot.mouseMove(206, 87);//opening data transfer window
            robotactions.add("opening data transfer window");
            Thread.sleep(1000);
            robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
            Thread.sleep(200);
            robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
            Thread.sleep(200);

        } else {
            robotactions.add("Action Failed");
        }

        if (Validation_AttendanceM_Opened() == true) {
            robot.mouseMove(561, 167);//setting data transfer mode to usb connect
            robotactions.add("setting data transfer mode to usb connect");
            Thread.sleep(200);
            robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
            Thread.sleep(200);
            robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
            Thread.sleep(2000);

        } else {
            robotactions.add("Action Failed");
        }

        if (Validation_AttendanceM_Opened() == true) {
            robot.mouseMove(531, 602);//pressing connect to connect to machine
            robotactions.add("pressing connect to connect to machine");
            Thread.sleep(200);
            robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
            Thread.sleep(200);
            robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
            Thread.sleep(2000);

        } else {
            robotactions.add("Action Failed");
        }

        if (Validation_AttendanceM_Opened() == true) {
            robot.mouseMove(752, 593);//pressing download data button
            robotactions.add("pressing download data button");
            Thread.sleep(200);
            robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
            Thread.sleep(200);
            robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
            Thread.sleep(1500);

        } else {
            robotactions.add("Action Failed");
        }

        if (Validation_AttendanceM_Opened() == true) {
            //613 437 yes button coordinates
            //761 , 434 no button coordinates
            robot.mouseMove(613, 437);//pressing yes on process attendance window
            robotactions.add("pressing yes on process attendance window");
            Thread.sleep(200);
            robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
            Thread.sleep(200);
            robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
            Thread.sleep(2000);

        } else {
            robotactions.add("Action Failed");
        }

        if (Validation_AttendanceM_Opened() == true) {
            robot.mouseMove(718, 433);//pressing ok on data process complete
            robotactions.add("pressing ok on data process complete");
            Thread.sleep(200);
            robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
            Thread.sleep(200);
            robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
            Thread.sleep(1500);
            robot.keyPress(KeyEvent.VK_ALT);
            robot.keyPress(KeyEvent.VK_TAB);
            Thread.sleep(1000);
            robot.keyRelease(KeyEvent.VK_ALT);
            robot.keyRelease(KeyEvent.VK_TAB);
            Thread.sleep(1000);

        } else {
            robotactions.add("Action Failed");
        }


        /* robot.mouseMove(531, 602);//pressing connect to connect to machine
    Thread.sleep(200);
    robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
    Thread.sleep(200);
    robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
    Thread.sleep(200);*/
 /* robot.mouseMove(865, 595);//pressing close button
    Thread.sleep(200);
    robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
    Thread.sleep(200);
    robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
    Thread.sleep(200);
  
    robot.mouseMove(1346, 7);//closing application
    Thread.sleep(200);
    robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
    Thread.sleep(200);
    robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
    Thread.sleep(200);*/
 /* robot.keyPress(KeyEvent.VK_WINDOWS);
   robot.keyPress(KeyEvent.VK_D);
   Thread.sleep(500);
   robot.keyRelease(KeyEvent.VK_WINDOWS);
   robot.keyRelease(KeyEvent.VK_D);*/
        if (Validation_AttendanceM_Opened() == true) {
            robot.mouseMove(865, 595);//pressing close button
            robotactions.add("pressing close button");
            Thread.sleep(200);
            robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
            Thread.sleep(200);
            robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
            Thread.sleep(200);
            Thread.sleep(1000);

        } else {
            robotactions.add("Action Failed");
        }

        if (Validation_AttendanceM_Opened() == true) {
            robot.mouseMove(338, 88);//clicking on data processing
            robotactions.add("clicking on data processing");
            Thread.sleep(200);
            robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
            Thread.sleep(200);
            robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
            Thread.sleep(200);
            Thread.sleep(1000);

        } else {
            robotactions.add("Action Failed");
        }

        if (Validation_AttendanceM_Opened() == true) {
            robot.mouseMove(799, 230);//clicking on process button
            robotactions.add("clicking on process button");
            Thread.sleep(200);
            robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
            Thread.sleep(200);
            robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
            Thread.sleep(3000);

        } else {
            robotactions.add("Action Failed");
        }

        if (Validation_AttendanceM_Opened() == true) {
            robot.mouseMove(698, 433);//clicking on ok button after processing
            robotactions.add("clicking on ok button after processing");
            Thread.sleep(200);
            robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
            Thread.sleep(200);
            robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
            Thread.sleep(200);

        } else {
            robotactions.add("Action Failed");
        }
        
        if (Validation_AttendanceM_Opened() == true) {
            Thread.sleep(1000);
            robot.mouseMove(1011, 229);//clicking close button after data processing
            robotactions.add("clicking close button after data processing");
            Thread.sleep(200);
            robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
            Thread.sleep(200);
            robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
            Thread.sleep(200);

        } else {
            robotactions.add("Action Failed");
        }

        if (Validation_AttendanceM_Opened() == true) {
            Thread.sleep(1000);
            robot.mouseMove(436, 34);//clicking on administration tab on menu bar
            robotactions.add("clicking on administration tab on menu bar");
            Thread.sleep(200);
            robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
            Thread.sleep(200);
            robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
            Thread.sleep(200);

        } else {
            robotactions.add("Action Failed");
        }

        if (Validation_AttendanceM_Opened() == true) {
            Thread.sleep(1000);
            robot.mouseMove(466, 122);//clicking on database backup from administarion menu
            robotactions.add("clicking on database backup from administarion menu");
            Thread.sleep(200);
            robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
            Thread.sleep(200);
            robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
            Thread.sleep(200);

        } else {
            robotactions.add("Action Failed");
        }

        if (Validation_AttendanceM_Opened() == true) {
            Thread.sleep(1000);
            robot.mouseMove(567, 321);//clicking on update database button
            robotactions.add("clicking on update database button");
            Thread.sleep(200);
            robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
            Thread.sleep(200);
            robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
            Thread.sleep(3000);

        } else {
            robotactions.add("Action Failed");
        }

        if (Validation_AttendanceM_Opened() == true) {
            Thread.sleep(1000);
            robot.mouseMove(726, 433);//clicking ok button after updating database
            robotactions.add("clicking ok button after updating database");
            Thread.sleep(200);
            robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
            Thread.sleep(200);
            robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
            Thread.sleep(200);
            Thread.sleep(1000);

        } else {
            robotactions.add("Action Failed");
        }

        if (Validation_AttendanceM_Opened() == true) {
            robot.mouseMove(533, 539);//clicking close button after updating database
            robotactions.add("clicking close button after updating database");
            Thread.sleep(200);
            robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
            Thread.sleep(200);
            robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
            Thread.sleep(200);

        } else {
            robotactions.add("Action Failed");
        }

        if (Validation_AttendanceM_Opened() == true) {
            Thread.sleep(1000);
            robot.mouseMove(1342, 9);//closing application
            robotactions.add("closing application");
            Thread.sleep(200);
            robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
            Thread.sleep(200);
            robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
            Thread.sleep(200);

        } else {
            robotactions.add("Action Failed");
        }
        if (Validation_AttendanceM_Opened() == true) {
            robotactions.add("Process Completed");
        } else {
            robotactions.add("Action Failed");
        }

   /*Thread.sleep(500);
   robot.keyPress(KeyEvent.VK_WINDOWS);
   robot.keyPress(KeyEvent.VK_SHIFT);
   robot.keyPress(KeyEvent.VK_M);
 
   Thread.sleep(500);
   robot.keyRelease(KeyEvent.VK_WINDOWS);
   robot.keyRelease(KeyEvent.VK_M);
   robot.keyRelease(KeyEvent.VK_SHIFT);*/
        System.out.println("");
        System.out.println(" Required Actions Arraylist:" + actions);
        System.out.println(" Required Actions Arraylist Size:" + actions.size());
        System.out.println("");

        System.out.println("Robot Actions Arraylist:" + robotactions);
        System.out.println("Robot Actions Arraylist Size:" + robotactions.size());
        System.out.println("");

        System.out.println("Native Actions Arraylist" + nativeactions);
        System.out.println("Native Actions Arraylist Size" + nativeactions.size());
        System.out.println("");

        Thread.sleep(500);

    }

}
