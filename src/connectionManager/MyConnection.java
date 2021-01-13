package connectionManager;

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
/**
 *
 * @author Khushbu
 */
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.logging.Level;
import java.util.logging.Logger;

public class MyConnection {

    public static Connection connection = null;

    /*   public static void main(String args[]) {
        getConnection("mytest");
        System.out.println("insert:" + insertData("insert into mytest.top_portfolio_websites (website_owner_name,website_url) values ('akshar','akshar.com')"));
        ResultSet rs = getResultSet("select * from mytest.top_portfolio_websites");
        try {
            while (rs.next()) {
                System.out.println(rs.getString("website_owner_name") + " : " + rs.getString("website_url"));
            }
        } catch (SQLException ex) {
            Logger.getLogger(MyConnection.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
     */
    public static Connection getConnection(String databaseName) {
        if (connection == null) {
            try {
                Class.forName("com.mysql.cj.jdbc.Driver");
                connection = (Connection) DriverManager.getConnection("jdbc:mysql://localhost:3306/" + databaseName, "root", "root");
                System.out.println("Connetion Established...");
            } catch (SQLException ex) {
                Logger.getLogger(MyConnection.class.getName()).log(Level.SEVERE, null, ex);
            } catch (ClassNotFoundException ex) {
                Logger.getLogger(MyConnection.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
        return connection;
    }

    public static Connection getArabicConnection(String databaseName) {
        if (connection == null) {
            try {
                Class.forName("com.mysql.jdbc.Driver");
                String unicode = "?useUnicode=yes&characterEncoding=UTF-8";
                connection = (Connection) DriverManager.getConnection("jdbc:mysql://localhost:3306/" + databaseName + unicode, "root", "root");
                System.out.println("Connetion Established...");
            } catch (SQLException ex) {
                Logger.getLogger(MyConnection.class.getName()).log(Level.SEVERE, null, ex);
            } catch (ClassNotFoundException ex) {
                Logger.getLogger(MyConnection.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
        return connection;
    }

    public static boolean insertData(String insertQuery) {
        PreparedStatement preparedStatement;
        boolean result = false;

        try {
            preparedStatement = connection.prepareStatement(insertQuery);
            int isInserted = preparedStatement.executeUpdate();
            if (isInserted == 1) {
                result = true;
            }
            //        System.out.println("res:" + result);
        } catch (SQLException ex) {
            if (ex.toString().contains("Duplicate entry")) {
                //---duplicate link...
            } else {
                System.out.println("Query:" + insertQuery);
                Logger.getLogger(MyConnection.class.getName()).log(Level.SEVERE, null, ex);
            }
            // System.out.println("Exception:"+ex);
        }

        return result;
    }

    public static ResultSet getResultSet(String query) {
        ResultSet rs = null;
        try {
            PreparedStatement preparedStatement = connection.prepareStatement(query);
            rs = preparedStatement.executeQuery();
        } catch (SQLException ex) {
            Logger.getLogger(MyConnection.class.getName()).log(Level.SEVERE, null, ex);
        }
        return rs;
    }

}
