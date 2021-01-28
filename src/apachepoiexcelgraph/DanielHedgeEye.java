/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package apachepoiexcelgraph;

import connectionManager.MyConnection;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedCondition;
import org.openqa.selenium.support.ui.WebDriverWait;

/**
 *
 * @author Khushbu
 */
public class DanielHedgeEye {

    static String filePath = "\\output\\HedgeEyeData.xlsx";
    static final int DATE_ROW = 1;
    static final int HEADER_ROW = 2;
    static int BUY_COL = 1;
    static int M1_COL = 2;
    static int COMP1_COL = 3;
    static int SELL_COL = 4;
    static int M2_COL = 5;
    static int COMP2_COL = 6;
    static int PREV_COL = 7;
    static int TOTAL_COL = 7;
    static String chromePath = "C:\\Users\\Khushbu\\Downloads\\chromedriver_win32(2)\\chromedriver.exe";
    static String dirPath = "";

    public static void main(String[] args) {
        if (args.length < 2) {
            System.out.println("Please check command line arguments...Need two paths");
            return;
        }
        dirPath = args[0];
        filePath = dirPath + filePath;
        chromePath = args[1];
        startScrape();
    }

    private static void startScrape() {
        login();
    }

    private static void login() {
        String loginURL = "https://accounts.hedgeye.com/users/sign_in";
        // String url="https://app.hedgeye.com/feed_items/84347-may-12-2020?with_category=33-risk-ranges";
        System.setProperty("webdriver.chrome.driver", chromePath);
        ChromeOptions options = new ChromeOptions();
        options.addArguments("--headless");
        ChromeDriver driver = new ChromeDriver(options);
        boolean hasNextPage = false;

        driver.get(loginURL);
        waitForJSandJQueryToLoad(driver);

        WebElement form = driver.findElementById("new_user");
        /* form.findElement(By.id("user_email")).sendKeys("bernd.sischka@gmail.com");
        form.findElement(By.id("user_password")).sendKeys("3qsd6gTQtXqw2&8EH#G4pv9x%H%@iDsyK1E$03ibpNmV^dwhkyAjvbJ072Ktgd$$gWmtE7#Y0QwYWNT");*/
        form.findElement(By.id("user_email")).sendKeys("dcooper@paradigmpmo.com");
        form.findElement(By.id("user_password")).sendKeys("H4mzapassword");
        form.findElement(By.id("se-be-login-submit")).click();

        waitForJSandJQueryToLoad(driver);

        getTodaysData(driver);

        generateExcelFile();

        DanielHedgeEyeGraph.GenerateGraph(dirPath);

        try {
            System.out.println("Sleeping..");
            Thread.sleep(10000);
        } catch (InterruptedException ex) {
            Logger.getLogger(DanielHedgeEye.class.getName()).log(Level.SEVERE, null, ex);
        }

        logout(driver);
        System.out.println("Logged out...");
    }

    private static void logout(ChromeDriver driver) {
        String logoutURL = "https://accounts.hedgeye.com/users/sign_out";
        driver.get(logoutURL);
        driver.close();
    }

    private static void getTodaysData(ChromeDriver driver) {
        try {
            //System.out.println(""+driver.getPageSource());
            driver.get("https://app.hedgeye.com/feed_items/all?page=1&amp;with_category=33-risk-ranges");
            //   driver.get("https://app.hedgeye.com/feed_items/92222-november-27-2020?with_category=33-risk-ranges");
            waitForJSandJQueryToLoad(driver);
            Document doc = Jsoup.parse(driver.getPageSource());
            Element table = doc.getElementsByClass("dtr-table").first();

            //get all index from database - check with today's index and add if any new index found
            MyConnection.getConnection("hedgeeye_tool");
            String selectQ = "SELECT index_id,index_name FROM hedgeeye_tool.index_master;";
            ResultSet rs = MyConnection.getResultSet(selectQ);
            while (rs.next()) {
                boolean matchFound = false;
                boolean isFirstRow = true;

                String value = "";
                for (Element tr : table.getElementsByTag("tr")) {
                    if (isFirstRow) {
                        isFirstRow = false;
                        continue;
                    }
                    String index = tr.getElementsByTag("td").get(0).getElementsByTag("strong").text();
                    if (index.contains("(")) {
                        index = StringUtils.substringBefore(index, "(").trim();
                    }
                    value = index;
                    String dbIndex = rs.getString("index_name");
                    if (dbIndex.equalsIgnoreCase(index)) {
                        matchFound = true;
                        break;
                    }
                }
                if (!matchFound) {
                    //insert in DB
                    System.out.println("New index found..." + value);
                    String insertQ = "insert into hedgeeye_tool.index_master (index_name) values ('" + value + "')";
                    MyConnection.getConnection("hedgeeye_tool");
                    MyConnection.insertData(insertQ);
                }
            }

            //add data to data_master
            String date = doc.getElementsByClass("article__header").first().text();
            SimpleDateFormat smt = new SimpleDateFormat("MMMM dd, yyyy");
            Date parse = smt.parse(date);
            smt = new SimpleDateFormat("yyyy-MM-dd");
            String feed_date = smt.format(parse);
            rs = MyConnection.getResultSet(selectQ);
            while (rs.next()) {
                boolean isFirstRow = true;
                for (Element tr : table.getElementsByTag("tr")) {
                    if (isFirstRow) {
                        isFirstRow = false;
                        continue;
                    }
                    String index = tr.getElementsByTag("td").get(0).getElementsByTag("strong").text();
                    if (index.contains("(")) {
                        index = StringUtils.substringBefore(index, "(").trim();
                    }
                    String dbIndex = rs.getString("index_name");
                    if (dbIndex.equalsIgnoreCase(index)) {
                        String buyTrade = tr.getElementsByTag("td").get(1).text();
                        String sellTrade = tr.getElementsByTag("td").get(2).text();
                        String prevclose = tr.getElementsByTag("td").get(3).text();
                        String trend = tr.attr("class");

                        //  System.out.println("" + index + ";" + buyTrade + ";" + sellTrade + ";" + prevclose);
                        String insertQ = "INSERT INTO `hedgeeye_tool`.`data_master`\n"
                                + "(\n"
                                + "`index_id`,\n"
                                + "`date`,\n"
                                + "`trend`,\n"
                                + "`buy`,\n"
                                + "`sell`,\n"
                                + "`prev_close`)\n"
                                + "VALUES\n"
                                + "("
                                + rs.getInt("index_id")
                                + ",'" + feed_date + "',"
                                + "'" + trend + "',"
                                + buyTrade.replaceAll("[^0-9.]", "") + ","
                                + sellTrade.replaceAll("[^0-9.]", "") + ","
                                + prevclose.replaceAll("[^0-9.]", "")
                                + ");";
                        MyConnection.getConnection("hedgeeye_tool");
                        MyConnection.insertData(insertQ);
                        //       System.out.println("INserted::" + index);
                        break;
                    }
                }
            }

        } catch (SQLException ex) {
            Logger.getLogger(DanielHedgeEye.class.getName()).log(Level.SEVERE, null, ex);
        } catch (ParseException ex) {
            Logger.getLogger(DanielHedgeEye.class.getName()).log(Level.SEVERE, null, ex);
        }

    }

    private static void generateExcelFile() {

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("sheet1");
        int rowno = 0;
        //  int cno = 1;
        int datecolno = 1;
        int headrecolno = 1;
        int totalIndex = 3;

        CellStyle greenStyle = workbook.createCellStyle();
        // Setting Background color  
        greenStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
        greenStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        CellStyle whiteStyle = workbook.createCellStyle();
        // Setting Background color  
        whiteStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        whiteStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        CellStyle redStyle = workbook.createCellStyle();
        // Setting Background color  
        redStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
        redStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        CellStyle greyStyle = workbook.createCellStyle();
        // Setting Background color  
        greyStyle.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
        greyStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        //create excel file format
        String selectQ = "SELECT index_name FROM hedgeeye_tool.index_master order by index_id;";
        MyConnection.getConnection("hedgeeye_tool");
        ResultSet rs = MyConnection.getResultSet(selectQ);
        int r = 3;
        //create date and header row
        Row row = sheet.createRow(DATE_ROW);
        row = sheet.createRow(HEADER_ROW);
        HashMap<String, String> indexMap = new HashMap();
        try {

            while (rs.next()) {
                row = sheet.createRow(r++);
                String index = rs.getString("index_name");
                row.createCell(0).setCellValue(index);
                indexMap.put("" + (r - 1), index);
                totalIndex++;
            }
        } catch (SQLException ex) {
            Logger.getLogger(DanielHedgeEye.class.getName()).log(Level.SEVERE, null, ex);
        }
        Date sdate = new Date(System.currentTimeMillis());
        Calendar c = Calendar.getInstance();
        c.add(Calendar.DAY_OF_MONTH, -30);
        Date dtEnd = c.getTime();
        System.out.println("today:" + sdate + " end date:" + dtEnd);
        DateFormat frmt = new SimpleDateFormat("yyyy-MM-dd");
        // String format = frmt.format(dtEnd);
        Calendar cp = Calendar.getInstance();
        cp.setTime(sdate);
        //This is to get next day in loop
        Calendar today = Calendar.getInstance();
        today.setTime(sdate);
        SimpleDateFormat smt = new SimpleDateFormat("MMMM d, yyyy");
        while (sdate.compareTo(dtEnd) > 0) {
            //continue if weekend
            System.out.println("Date:" + sdate);
            today.setTime(sdate);
            if (today.get(Calendar.DAY_OF_WEEK) == 7 || today.get(Calendar.DAY_OF_WEEK) == 0) {
                System.out.println("weeked..");
                today.add(Calendar.DAY_OF_MONTH, -1);
                sdate = today.getTime();
                continue;
            }

            //add current date data for all index
            if (dateHasRecords(frmt.format(sdate))) {
                //add date in second row. 
                if (sheet.getRow(DATE_ROW).getCell(datecolno) == null) {
                    sheet.getRow(DATE_ROW).createCell(datecolno).setCellValue(smt.format(sdate));
                    datecolno = datecolno + 5;
                }
                if (sheet.getRow(HEADER_ROW).getCell(headrecolno) == null) {
                    sheet.getRow(HEADER_ROW).createCell(headrecolno++).setCellValue("BUY");
                    sheet.getRow(HEADER_ROW).createCell(headrecolno++).setCellValue("Movement");
                    sheet.getRow(HEADER_ROW).createCell(headrecolno++).setCellValue("Prev Buy - BUY");
                    sheet.getRow(HEADER_ROW).createCell(headrecolno++).setCellValue("SELL");
                    sheet.getRow(HEADER_ROW).createCell(headrecolno++).setCellValue("Movement");
                    sheet.getRow(HEADER_ROW).createCell(headrecolno++).setCellValue("Prev Sell - Sell");
                    //.getRow(HEADER_ROW).createCell(headrecolno++).setCellValue("BUY");
                    sheet.getRow(HEADER_ROW).createCell(headrecolno++).setCellValue("PREV.CLOSE");
                }

                //Find prev record date
                boolean isPrev = false;
                int prevScanCount = 1;
                Date pdate = cp.getTime();
                try {

                    cp.setTime(today.getTime());

                    while (!isPrev && prevScanCount < 10) {
                        cp.add(Calendar.DAY_OF_MONTH, -1);
                        pdate = cp.getTime();
                        String str = smt.format(pdate);
                        pdate = smt.parse(str);

                        System.out.println("Finding previous day " + pdate + " : " + str);
                        isPrev = getPreviousDay(frmt.format(pdate));
                        cp.setTime(pdate);
                        prevScanCount++;
                    }

                } catch (ParseException ex) {
                    Logger.getLogger(DanielHedgeEye.class.getName()).log(Level.SEVERE, null, ex);
                }

                for (rowno = 3; rowno < totalIndex; rowno++) {
                    String indexAtRow = indexMap.get("" + rowno);
                    //   System.out.println("--->" + indexAtRow);
                    if (indexAtRow != null) {
                        String Q1 = "SELECT trend,buy,sell,prev_close FROM hedgeeye_tool.data_master d, hedgeeye_tool.index_master l where "
                                + "d.index_id=l.index_id and `date`='" + frmt.format(sdate) + "' and `index_name`='" + indexAtRow + "';";
                        String Q2 = "SELECT trend,buy,sell,prev_close FROM hedgeeye_tool.data_master d, hedgeeye_tool.index_master l where "
                                + "d.index_id=l.index_id and `date`='" + frmt.format(pdate) + "' and `index_name`='" + indexAtRow + "';";
                        //     System.out.println("" + Q1);
                        MyConnection.getConnection("hedgeeye_tool");
                        ResultSet rsCurrent = MyConnection.getResultSet(Q1);
                        ResultSet rsPrev = MyConnection.getResultSet(Q2);
                        try {
                            if (rsCurrent.next()) {
                                //  System.out.println("row no..." + rowno + "Adding data at column..." + BUY_COL);
                                //  System.out.println("s:"+SELL_COL+";P:"+PREV_COL+";M1:"+M1_COL+";M2:"+M2_COL);
                                double buy_trade = Double.parseDouble(rsCurrent.getString("buy"));
                                double sale_trade = Double.parseDouble(rsCurrent.getString("sell"));
                                double prev_close = Double.parseDouble(rsCurrent.getString("prev_close"));
                                String trend = rsCurrent.getString("trend");
                                Row rw = sheet.getRow(rowno);
                                Cell cell1 = rw.createCell(BUY_COL);
                                cell1.setCellValue(buy_trade);
                                Cell cell2 = rw.createCell(SELL_COL);
                                cell2.setCellValue(sale_trade);
                                rw.createCell(PREV_COL).setCellValue(prev_close);

                                if (trend.equalsIgnoreCase("bearish")) {
                                    cell1.setCellStyle(redStyle);
                                    cell2.setCellStyle(redStyle);
                                } else if (trend.equalsIgnoreCase("bullish")) {
                                    cell1.setCellStyle(greenStyle);
                                    cell2.setCellStyle(greenStyle);
                                } else if (trend.equalsIgnoreCase("neutral")) {
                                    cell1.setCellStyle(greyStyle);
                                    cell2.setCellStyle(greyStyle);
                                }
                                if (rsPrev.next()) {
                                    double buy_trade1 = Double.parseDouble(rsPrev.getString("buy"));
                                    double sale_trade1 = Double.parseDouble(rsPrev.getString("sell"));

                                    Cell cl = rw.createCell(COMP1_COL);
                                    cl.setCellValue(buy_trade1 - buy_trade);

                                    cl = rw.createCell(COMP2_COL);
                                    cl.setCellValue(sale_trade1 - sale_trade);

                                    if (buy_trade > buy_trade1) {
                                        Cell cell = rw.createCell(M1_COL);
                                        cell.setCellValue("UP");
                                        cell.setCellStyle(greenStyle);
                                    } else if (buy_trade < buy_trade1) {
                                        Cell cell = rw.createCell(M1_COL);
                                        cell.setCellValue("DOWN");
                                        cell.setCellStyle(redStyle);
                                    } else {
                                        Cell cell = rw.getCell(BUY_COL);
                                        cell.setCellValue("NA");
                                        cell.setCellStyle(whiteStyle);
                                    }
                                    if (sale_trade > sale_trade1) {
                                        Cell cell = rw.createCell(M2_COL);
                                        cell.setCellValue("UP");
                                        cell.setCellStyle(greenStyle);
                                    } else if (sale_trade < sale_trade1) {
                                        Cell cell = rw.createCell(M2_COL);
                                        cell.setCellValue("DOWN");
                                        cell.setCellStyle(redStyle);
                                    } else {
                                        Cell cell = rw.getCell(SELL_COL);
                                        cell.setCellValue("NA");
                                        cell.setCellStyle(whiteStyle);
                                    }
                                }
                            }
                        } catch (SQLException ex) {
                            Logger.getLogger(DanielHedgeEye.class.getName()).log(Level.SEVERE, null, ex);
                        }
                    }
                }
                BUY_COL = BUY_COL + TOTAL_COL;
                SELL_COL = SELL_COL + TOTAL_COL;
                PREV_COL = PREV_COL + TOTAL_COL;
                M1_COL = M1_COL + TOTAL_COL;
                M2_COL = M2_COL + TOTAL_COL;
                COMP1_COL = COMP1_COL + TOTAL_COL;
                COMP2_COL = COMP2_COL + TOTAL_COL;
                // System.out.println("Column increased>>>BUY col:" + BUY_COL);
            }
            //get previous date

            today.add(Calendar.DAY_OF_YEAR, -1);
            sdate = today.getTime();
            //  System.out.println("Prev->" + frmt.format(sdate));

        }

        ///write output to file
        FileOutputStream fileOutputStream = null;
        File file = new File(filePath);
        if (file.exists()) {
            file.delete();
        }
        try {
            file.createNewFile();
        } catch (IOException ex) {
            Logger.getLogger(DanielHedgeEye.class.getName()).log(Level.SEVERE, null, ex);
        }
        // String filePath = "E:\\output\\TheWholeSaleSuppliers\\" + fileName + ".xlsx";
        try {
            fileOutputStream = new FileOutputStream(file);
            workbook.write(fileOutputStream);
            workbook.close();
            fileOutputStream.flush();
            fileOutputStream.close();
            System.out.println("Get your file at:" + filePath);

        } catch (FileNotFoundException ex) {
            Logger.getLogger(DanielHedgeEye.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(DanielHedgeEye.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    private static boolean dateHasRecords(String sdate) {
        String selectQ = "SELECT * FROM hedgeeye_tool.data_master where `date`='" + sdate + "'; ";
        MyConnection.getConnection("hedgeeye_tool");
        ResultSet rs = MyConnection.getResultSet(selectQ);
        try {
            if (rs.next()) {
                return true;
            }
        } catch (SQLException ex) {
            Logger.getLogger(DanielHedgeEye.class.getName()).log(Level.SEVERE, null, ex);
        }
        return false;
    }

    private static boolean getPreviousDay(String str) {
        String selectQ = "SELECT * FROM hedgeeye_tool.data_master where `date`='" + str + "'; ";
        //  System.out.println("" + selectQ);
        MyConnection.getConnection("hedgeeye_tool");
        ResultSet rs = MyConnection.getResultSet(selectQ);
        try {
            if (rs.next()) {
                return true;
            }
        } catch (SQLException ex) {
            Logger.getLogger(DanielHedgeEye.class.getName()).log(Level.SEVERE, null, ex);
        }
        return false;
    }

    public static boolean waitForJSandJQueryToLoad(ChromeDriver driver) {

        WebDriverWait wait = new WebDriverWait(driver, 30);

        // wait for jQuery to load
        ExpectedCondition<Boolean> jQueryLoad = new ExpectedCondition<Boolean>() {
            @Override
            public Boolean apply(WebDriver driver) {
                try {
                    // return ((Long) ((JavascriptExecutor) getDriver()).executeScript("return jQuery.active") == 0);
                    return ((JavascriptExecutor) driver).executeScript("return jQuery.active == 0").equals(true);
                } catch (Exception e) {
                    // no jQuery present
                    return true;
                }
            }
        };

        // wait for Javascript to load
        ExpectedCondition<Boolean> jsLoad = new ExpectedCondition<Boolean>() {
            @Override
            public Boolean apply(WebDriver driver) {
                // return ((JavascriptExecutor) getDriver()).executeScript("return document.readyState")
                //        .toString().equals("complete");
                return ((JavascriptExecutor) driver).executeScript("return document.readyState").equals("complete");
            }
        };

        return wait.until(jQueryLoad) && wait.until(jsLoad);
    }
}
