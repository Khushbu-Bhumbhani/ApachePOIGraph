/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package apachepoiexcelgraph;

import connectionManager.MyConnection;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.YearMonth;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.LegendPosition;
import org.apache.poi.xddf.usermodel.chart.MarkerStyle;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFChartLegend;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFLineChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Khushbu
 */
public class DanielHedgeEyeGraph {

    public static void main(String[] args) {
        try {
            XSSFWorkbook wb = new XSSFWorkbook();
            String sheetName = "Sheet1";
            FileOutputStream fileOut = null;
            String selectIndex = "SELECT index_name FROM hedgeeye_tool.index_master;";
            MyConnection.getConnection("hedgeeye_tool");
            ResultSet rsIndex = MyConnection.getResultSet(selectIndex);

            // String indexName = "UST10Y";
            int rowCout = 5;
            int colCount = 0;
            while (rsIndex.next()) {
                String indexName = rsIndex.getString("index_name");
                System.out.println("Index:->" + indexName);
                XSSFSheet sheet = wb.createSheet(indexName.replace("/", "_"));
                String selectQ = "SELECT * FROM hedgeeye_tool.data_master d, index_master i where d.index_id=i.index_id and index_name='" + indexName + "';";
                Double buyData[] = new Double[31];
                Double saleData1[] = new Double[31];
                //Double trendChange[] = new Double[31];

                ResultSet rs = MyConnection.getResultSet(selectQ);
                SimpleDateFormat smt = new SimpleDateFormat("yyyy-MM-dd");

                for (int i = 0; i <= 30; i++) {
                    buyData[i] = 0.0;
                    saleData1[i] = 0.0;
                    // trendChange[i] = 0.0;
                }
                //  int datCount = 0;
                //  String defaultTrend = "bullish";
                while (rs.next()) {
                    double buy = rs.getDouble("buy");
                    double sale = rs.getDouble("sell");
                    //  String trend = rs.getString("trend");
                    String date = rs.getString("date");
                    Date dt = smt.parse(date);
                    int arrayIndex = dt.getDate() - 1;
                    buyData[arrayIndex] = buy;
                    saleData1[arrayIndex] = sale;
                    /*  if (!trend.equalsIgnoreCase(defaultTrend)) {
                        trendChange[arrayIndex] = 1.0;
                        defaultTrend = trend;
                    }
                    datCount++;*/
                }

                /* double sum = 0;
                for (double value : buyData) {
                    sum += value;
                }
                 double buyDataAvg = sum / datCount;
            for (int i = 0; i <= 30; i++) {
                {
                    if (trendChange[i] == 1.0) {
                        trendChange[i] = buyDataAvg;
                    }
                }
            }*/
 /* for (double value : saleData1) {
                sum += value;
            }
            double saleDataAvg = sum / datCount;*/
                //Create a canvas
                XSSFDrawing drawing = sheet.createDrawingPatriarch();
                //The first four default 0, [0,5]: start from 0 column and 5 rows; [7,26]: width 7 cells, 26 expands down to 26 rows
                //Default width (14-8)*12
                XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, 5, 7, 26);
                //   XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, colCount, rowCout, 7, 26);
                //Create a chart object
                XSSFChart chart = drawing.createChart(anchor);
                //Title
                chart.setTitleText(indexName);
                //Title overwrite
                chart.setTitleOverlay(false);

                //Legend position
                XDDFChartLegend legend = chart.getOrAddLegend();
                legend.setPosition(LegendPosition.TOP);

                //Classification axis (X axis), title position
                XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
                bottomAxis.setTitle("Date");
                //Value (Y axis) axis, title position
                XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
                leftAxis.setTitle("Values");

                String dates[] = new String[31];
                // Get the number of days in that month
                Calendar calendar = Calendar.getInstance();
                YearMonth yearMonthObject = YearMonth.of(calendar.get(Calendar.YEAR), calendar.get(Calendar.MONTH) + 1);
                int daysInMonth = yearMonthObject.lengthOfMonth(); //28

                System.out.println("Days in moth:" + daysInMonth);
                for (int i = 1, j = 0; i <= daysInMonth; i++, j++) {
                    if (calendar.get(Calendar.DAY_OF_WEEK) != 7 && calendar.get(Calendar.DAY_OF_WEEK) != 0) {
                        dates[j] = i + "/" + (calendar.get(Calendar.MONTH) + 1);
                    }
                    // testData[j] = i;
                    // testData1[j] = 100 - i;
                }
                System.out.println("First date:" + dates[1]);

                XDDFCategoryDataSource countries = XDDFDataSourcesFactory.fromArray(dates);
                //Data 1, cell range position [1, 0] to [1, 6]
                //    XDDFNumericalDataSource<Double> area = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 1, 0, 6));
                XDDFNumericalDataSource<Double> buyDataSource = XDDFDataSourcesFactory.fromArray(buyData);

                //Data 1, cell range position [2, 0] to [2, 6]
                XDDFNumericalDataSource<Double> saleDataSource = XDDFDataSourcesFactory.fromArray(saleData1);
                //    XDDFNumericalDataSource<Double> trendDataSource = XDDFDataSourcesFactory.fromArray(trendChange);

                //LINE: line chart,
                XDDFLineChartData data = (XDDFLineChartData) chart.createData(ChartTypes.LINE, bottomAxis, leftAxis);

                //Chart load data, broken line 1
                XDDFLineChartData.Series series1 = (XDDFLineChartData.Series) data.addSeries(countries, buyDataSource);
                //Line legend title
                series1.setTitle("Buy", null);
                //Straight
                series1.setSmooth(false);
                //Set the mark size
                series1.setMarkerSize((short) 6);
                //Set the mark style, stars
                series1.setMarkerStyle(MarkerStyle.STAR);

                //Chart load data, broken line 2
                XDDFLineChartData.Series series2 = (XDDFLineChartData.Series) data.addSeries(countries, saleDataSource);
                //Line legend title
                series2.setTitle("Sale", null);
                //Curve
                series2.setSmooth(true);
                //Set the mark size
                series2.setMarkerSize((short) 6);
                //Set the mark style, square
                series2.setMarkerStyle(MarkerStyle.CIRCLE);
                /*     
              //Chart load data, broken line 3
            XDDFLineChartData.Series series3 = (XDDFLineChartData.Series) data.addSeries(countries, trendDataSource);
            //Line legend title
            series3.setTitle("Trend", null);
            //Curve
            series3.setSmooth(true);
            //Set the mark size
            series3.setMarkerSize((short) 6);
            //Set the mark style, square
            series3.setMarkerStyle(MarkerStyle.SQUARE);*/

                //Draw
                chart.plot(data);

                rowCout = rowCout + 30;
            }
            // Write output to excel file
            String filename = "E:\\DanielGraphTest.xlsx";
            fileOut = new FileOutputStream(filename);
            wb.write(fileOut);
            System.out.println("->" + filename);
        } catch (FileNotFoundException ex) {
            Logger.getLogger(DanielHedgeEyeGraph.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(DanielHedgeEyeGraph.class.getName()).log(Level.SEVERE, null, ex);
        } catch (SQLException ex) {
            Logger.getLogger(DanielHedgeEyeGraph.class.getName()).log(Level.SEVERE, null, ex);
        } catch (ParseException ex) {
            Logger.getLogger(DanielHedgeEyeGraph.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

}
