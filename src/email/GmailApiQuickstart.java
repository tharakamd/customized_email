/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package email;

/**
 *
 * @author tharaka
 */
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.auth.oauth2.GoogleCredential;
import com.google.api.client.googleapis.auth.oauth2.GoogleOAuthConstants;
import com.google.api.client.googleapis.auth.oauth2.GoogleTokenResponse;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.services.gmail.Gmail;
import com.google.api.services.gmail.model.Thread;
import com.google.api.services.gmail.model.ListThreadsResponse;

import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.Arrays;
import java.util.List;

import com.google.api.client.repackaged.org.apache.commons.codec.binary.Base64;
import com.google.api.services.gmail.Gmail;
import com.google.api.services.gmail.model.Draft;
import com.google.api.services.gmail.model.Message;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.Iterator;
import java.util.Properties;
import java.util.logging.Level;
import java.util.logging.Logger;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.Session;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class GmailApiQuickstart {

    // Check https://developers.google.com/gmail/api/auth/scopes for all available scopes
    private static final String SCOPE = "https://www.googleapis.com/auth/gmail.compose";
    private static final String APP_NAME = "Gmail API Quickstart";
    // Email address of the user, or "me" can be used to represent the currently authorized user.
    private static final String USER = "me";
    // Path to the client_secret.json file downloaded from the Developer Console
    private static final String CLIENT_SECRET_PATH = "client_secret_585909283629-g3nbr8ohc7duicsqn21s5lntt0hpk1o9.apps.googleusercontent.com.json";

    private static GoogleClientSecrets clientSecrets;

    public static void main(String[] args) throws IOException {
        ArrayList<String> data = readExcel();
        HttpTransport httpTransport = new NetHttpTransport();
        JsonFactory jsonFactory = new JacksonFactory();

        clientSecrets = GoogleClientSecrets.load(jsonFactory, new FileReader(CLIENT_SECRET_PATH));

        // Allow user to authorize via url.
        GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
                httpTransport, jsonFactory, clientSecrets, Arrays.asList(SCOPE))
                .setAccessType("online")
                .setApprovalPrompt("auto").build();

        String url = flow.newAuthorizationUrl().setRedirectUri(GoogleOAuthConstants.OOB_REDIRECT_URI)
                .build();
        System.out.println("Please open the following URL in your browser then type"
                + " the authorization code:\n" + url);

        // Read code entered by user.
        BufferedReader br = new BufferedReader(new InputStreamReader(System.in));
        String code = br.readLine();

        // Generate Credential using retrieved code.
        GoogleTokenResponse response = flow.newTokenRequest(code)
                .setRedirectUri(GoogleOAuthConstants.OOB_REDIRECT_URI).execute();
        GoogleCredential credential = new GoogleCredential()
                .setFromTokenResponse(response);

        // Create a new authorized Gmail API client
        Gmail service = new Gmail.Builder(httpTransport, jsonFactory, credential)
                .setApplicationName(APP_NAME).build();
        /*
         // Retrieve a page of Threads; max of 100 by default.
         ListThreadsResponse threadsResponse = service.users().threads().list(USER).execute();
         List<Thread> threads = threadsResponse.getThreads();

         // Print ID of each Thread.
         for (Thread thread : threads) {
         System.out.println("Thread ID: " + thread.getId());
         }
         */

        // html text
        String bodytext = "<p>Dear ###1###,</p>\n"
                + "        <h3 style=\"text-align: center; text-decoration: underline;\">Congratulations on Successfully Registering for IESL Idea Challenge 2015</h3>\n"
                + "        <p>We are pleased to inform you that you have successfully completed the registration process for IESL Idea Challenge 2015. You are now a registered team for the competition. Please feel free to address any concerns you may have regarding the competition to this email address. Further details and announcements would be made available on the official website and via this email.</p>\n"
                + "        <p>We congratulate you on your registration for the competition and wish you all the best going forward.</p>\n"
                + "        <p>Followings are your details:</p>\n"
                + "        <p>Name: ###1### </p>\n"
                + "        <p>TP:###2### </p>\n"
                + "        <p>School: ###3###</p>\n"
                + "        <p>Thank You.</p><br>\n"
                + "        <p>Best Regards,</p>\n"
                + "        <p>Organizing Committee,</p>\n"
                + "        <p>IDEA Challenge 2015</p>";
    //

        // calling the function
        createDraftEmail("tharakamd6@gmail.com", "dilanreader@gamil.com", "Hello Gmail", bodytext, service);
        sendEmail(data, "dilanreader@gamil.com", "Hello Gmail", bodytext, 3, service);

    }

    public static void createDraftEmail(String to, String from, String subject, String bodyText, Gmail service) {
        try {
            MimeMessage msg = createEmail(to, from, subject, bodyText);
            try {

                Draft draft = createDraft(service, USER, msg);

            } catch (IOException ex) {
                Logger.getLogger(GmailApiQuickstart.class.getName()).log(Level.SEVERE, null, ex);
            }

        } catch (MessagingException ex) {
            Logger.getLogger(GmailApiQuickstart.class.getName()).log(Level.SEVERE, null, ex);
        }

    }

    public static void sendEmail(ArrayList<String> data, String from, String subject, String bodyText, int number, Gmail service) {

        Iterator<String> it = data.iterator();
        int i;
        while (it.hasNext()) {
            i = 1;
            String tmp_body = bodyText;
            String to = it.next();
            for (int j = 1; j <= number; j++) {
                String text = "###" + String.valueOf(i) + "###";
                System.out.println(text);
                tmp_body = tmp_body.replaceAll(text, it.next());
                i++;

                
            }
            try {
                    MimeMessage msg = createEmail(to, from, subject, tmp_body);
                    try {

                        Draft draft = createDraft(service, USER, msg);

                    } catch (IOException ex) {
                        Logger.getLogger(GmailApiQuickstart.class.getName()).log(Level.SEVERE, null, ex);
                    }

                } catch (MessagingException ex) {
                    Logger.getLogger(GmailApiQuickstart.class.getName()).log(Level.SEVERE, null, ex);
                }


        }

    }

    // ...
    /**
     * Create draft email.
     *
     * @param service an authorized Gmail API instance
     * @param userId user's email address. The special value "me" can be used to
     * indicate the authenticated user
     * @param email the MimeMessage used as email within the draft
     * @return the created draft
     * @throws MessagingException
     * @throws IOException
     */
    public static Draft createDraft(Gmail service, String userId, MimeMessage email)
            throws MessagingException, IOException {
        Message message = createMessageWithEmail(email);
        Draft draft = new Draft();
        draft.setMessage(message);
        draft = service.users().drafts().create(userId, draft).execute();

        System.out.println("draft id: " + draft.getId());
        System.out.println(draft.toPrettyString());
        return draft;
    }

    /**
     * Create a message from an email
     *
     * @param email Email to be set to raw of message
     * @return a message containing a base64url encoded email
     * @throws IOException
     * @throws MessagingException
     */
    public static Message createMessageWithEmail(MimeMessage email)
            throws MessagingException, IOException {
        ByteArrayOutputStream bytes = new ByteArrayOutputStream();
        email.writeTo(bytes);
        String encodedEmail = Base64.encodeBase64URLSafeString(bytes.toByteArray());
        Message message = new Message();
        message.setRaw(encodedEmail);
        return message;
    }

    /**
     * Create a MimeMessage using the parameters provided.
     *
     * @param to email address of the receiver
     * @param from email address of the sender, the mailbox account
     * @param subject subject of the email
     * @param bodyText body text of the email
     * @return the MimeMessage to be used to send email
     * @throws MessagingException
     */
    public static MimeMessage createEmail(String to, String from, String subject,
            String bodyText) throws MessagingException {
        Properties props = new Properties();
        Session session = Session.getDefaultInstance(props, null);

        MimeMessage email = new MimeMessage(session);

        email.setFrom(new InternetAddress(from));
        email.addRecipient(javax.mail.Message.RecipientType.TO,
                new InternetAddress(to));
        email.setSubject(subject);
        // email.setText("<h1>Hello</h1>");

        Multipart mp = new MimeMultipart();
        MimeBodyPart htmlPart = new MimeBodyPart();
        htmlPart.setContent(bodyText, "text/html");
        mp.addBodyPart(htmlPart);
        email.setContent(mp);

        return email;
    }

    // ...
    public static ArrayList<String> readExcel() {
        ArrayList<String> reading = new ArrayList<>();
        Workbook wb;
        try {
            wb = WorkbookFactory.create(new File("Book1.xls"));
            Sheet sheet1 = wb.getSheetAt(0);
            for (Row row : sheet1) {
                for (Cell cell : row) {
                    CellReference cellRef = new CellReference(row.getRowNum(), cell.getColumnIndex());
                 //   System.out.print(cellRef.formatAsString());
                    //   System.out.print(" - ");

                    switch (cell.getCellType()) {
                        case Cell.CELL_TYPE_STRING:
                            // System.out.println(cell.getRichStringCellValue().getString());
                            reading.add(cell.getRichStringCellValue().getString());
                            break;
                        case Cell.CELL_TYPE_NUMERIC:
                            if (DateUtil.isCellDateFormatted(cell)) {
                                //System.out.println(cell.getDateCellValue());
                                reading.add(cell.getDateCellValue().toString());
                            } else {
                                //System.out.println(cell.getNumericCellValue());

                                reading.add(String.valueOf(cell.getNumericCellValue()));
                            }
                            break;
                        case Cell.CELL_TYPE_BOOLEAN:
                            //System.out.println(cell.getBooleanCellValue());
                            reading.add(String.valueOf(cell.getBooleanCellValue()));
                            break;
                        case Cell.CELL_TYPE_FORMULA:
                            //System.out.println(cell.getCellFormula());
                            reading.add(String.valueOf(cell.getCellFormula()));
                            break;
                        default:
                        // System.out.println();
                    }
                }
            }

        } catch (IOException ex) {
            Logger.getLogger(GmailApiQuickstart.class.getName()).log(Level.SEVERE, null, ex);
        } catch (InvalidFormatException ex) {
            Logger.getLogger(GmailApiQuickstart.class.getName()).log(Level.SEVERE, null, ex);
        }
        return reading;

    }

}
