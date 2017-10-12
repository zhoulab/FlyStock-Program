import java.text.ParseException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.io.File;

public class Query
{
    public static void main(String[] args) throws ParseException
    {
        // Arguments necessary to query
        if(args.length!=0){
        try
        {
            if(args[0].contains(" ")){
              String[] argument = (args[0]).split(" ");
              args = argument;
            }

            // Get current working directory
            String db_file_name_prefix = (System.getProperty("user.dir")+"\\database\\mydb");

            Connection con = null;

            // Load the HSQL Database Engine JDBC driver
            // hsqldb.jar should be in the class path or made part of the current jar
            Class.forName("org.hsqldb.jdbcDriver");

            // connect to the database.   This will load the db files and start the
            // database if it is not alread running.
            // db_file_name_prefix is used to open or create files that hold the state
            // of the db.
            // It can contain directory names relative to the
            // current working directory
            con = DriverManager.getConnection("jdbc:hsqldb:file:" + db_file_name_prefix, // filenames
                  "sa", // username if there is any
                    "");  // password if there is any

            Statement statement = con.createStatement();

            //QUERYING
            //Example: args[0] = Stock_ID, args[1] = G001.
            ResultSet rs = statement.executeQuery("SELECT \"Stock_ID\", \"Genotype\", \"Description\", \"Note\", \"Res_Person\", \"Flybase\", \"Project\" "
                                                +"FROM \"Fly_Stock\" WHERE ( UPPER ( \"" + args[0] + "\" ) LIKE UPPER ( '%" + args[1] + "%' ) )");
                                                //+"OR UPPER ( \"Genotype\" ) LIKE UPPER ( '%" + args[0] + "%' ) "
                                                //+"OR UPPER ( \"Description\" ) LIKE UPPER ( '%" + args[0] + "%' ) "
                                                //+"OR UPPER ( \"Res_Person\" ) LIKE UPPER ( '%" + args[0] + "%' ) ) "
                                                //+"OR UPPER ( \"Note\" ) LIKE UPPER ( '%" + args[0] + "%' )");

            //print the result set seperated by |||
            while (rs.next())
            {
                System.out.println(rs.getString("Stock_ID")
                + " ||| " + rs.getString("Genotype")
                + " ||| " + rs.getString("Description")
                + " ||| " + rs.getString("Note")
                + " ||| " + rs.getString("Res_Person")
                + " ||| " + rs.getString("Flybase")
                + " ||| " + rs.getString("Project"));
            }

            // Properly exit and close connections
            statement.close();
            con.close();
            rs.close();

        // Exception handling
        } catch (SQLException ex)
        {
            Logger.getLogger(Query.class.getName()).log(Level.SEVERE, null, ex);
            ex.printStackTrace();
        } catch (ClassNotFoundException ex)
        {
            Logger.getLogger(Query.class.getName()).log(Level.SEVERE, null, ex);
        }
        }
    }
}
