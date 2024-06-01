package projectfinalpackage;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class WordPermissionChecker {

    public static boolean hasWritePermission(String filePath) {
        try {
            // Try to open the file in write mode
            FileOutputStream out = new FileOutputStream(new File(filePath), true);
            out.close();
            // If no exception is thrown, it means write permission is granted
            return true;
        } catch (IOException e) {
            // If an IOException occurs, it means write permission is not granted
            return false;
        }
    }

}
