package com.aspose.spreadsheeteditor;

import com.aspose.cells.CellsException;
import com.aspose.cells.License;
import java.io.IOException;
import java.io.InputStream;
import java.util.Date;
import java.util.logging.Logger;

/**
 *
 * @author Saqib Masood
 */
public class AsposeLicense {

    private static final Logger LOGGER = Logger.getLogger(AsposeLicense.class.getName());

    public static final String FILE_NAME = "Aspose.Total.Java.lic";

    private AsposeLicense() {
    }

    public static void load() {
        Date expiry = License.getSubscriptionExpireDate();

        if (expiry != null) {
            LOGGER.info(String.format("Aspose License is valid upto: %s", License.getSubscriptionExpireDate()));
        }
    }

    static {
        try (InputStream i = AsposeLicense.class.getResourceAsStream(FILE_NAME)) {
            new License().setLicense(i);
        } catch (IOException x) {
            LOGGER.severe("Error occured while loading license");
            LOGGER.throwing(null, null, x);
        } catch (CellsException x) {
            LOGGER.severe("License error");
            LOGGER.throwing(null, null, x);
        }
    }
}
