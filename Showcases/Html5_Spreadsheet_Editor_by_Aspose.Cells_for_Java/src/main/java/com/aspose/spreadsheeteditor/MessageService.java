package com.aspose.spreadsheeteditor;

import java.util.logging.Logger;
import javax.enterprise.context.ApplicationScoped;
import javax.faces.application.FacesMessage;
import javax.faces.context.FacesContext;
import javax.inject.Named;
import org.primefaces.context.RequestContext;

/**
 *
 * @author Saqib Masood
 */
@Named(value = "msg")
@ApplicationScoped
public class MessageService {

    private static final Logger LOGGER = Logger.getLogger(MessageService.class.getName());

    public void sendMessage(String summary, String details) {
        LOGGER.info(String.format("%s: %s", summary, details));
        FacesContext.getCurrentInstance().addMessage(null, new FacesMessage(summary, details));
    }

    public void sendMessageDialog(String summary, String details) {
        LOGGER.info(String.format("%s: %s", summary, details));
        RequestContext.getCurrentInstance().showMessageInDialog(new FacesMessage(summary, details));
    }
}
