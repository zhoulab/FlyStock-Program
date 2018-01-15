import java.text.ParseException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.io.File;

public class Add
{
    public static void main(String[] args) throws ParseException
    {
      Connection con = null;
      Statement statement = null;
      ResultSet rs = null;

        try {
            // Get current working directory
            String db_file_name_prefix = (System.getProperty("user.dir")+"\\LabDB\\database\\mydb");
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

            statement = con.createStatement();

            //QUERYING
            //Example: args[0] = Stock_ID, args[1] = G001.
            String command = "INSERT INTO \"Fly_Stock\" (\"Stock_ID\", \"Genotype\", \"Description\", \"Note\", \"Res_Person\", \"Flybase\", \"Project\") VALUES (\'M018\', \'S\', \'N\', \'note\', \'res\', \'fly\', \'proj\')";

            statement.execute(command);

            //Writes
            /*    INSERT INTO "Fly_Stock" ("Stock_ID", "Genotype", "Description", "Note", "Res_Person", "Flybase", "Project")
    VALUES ('M500', 'S', 'N', 'note', 'res', 'fly', 'proj')*/

        // Exception handling
        }
         catch (SQLException ex)
        {
            Logger.getLogger(Query.class.getName()).log(Level.SEVERE, null, ex);
            ex.printStackTrace();
        } catch (ClassNotFoundException ex)
        {
            Logger.getLogger(Query.class.getName()).log(Level.SEVERE, null, ex);
        }
        // Properly exit and close connections
        finally {
          try { rs.close(); } catch (Exception e) { /* ignored */ }
          try { statement.close(); } catch (Exception e) { /* ignored */ }
          try { con.close(); } catch (Exception e) { /* ignored */ }
}
        }
    }
