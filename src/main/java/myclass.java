import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;

import com.google.api.services.gmail.GmailScopes;
import com.google.api.services.gmail.model.*;
import com.google.api.services.gmail.Gmail;

import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;



import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.*;
import com.google.api.services.sheets.v4.Sheets;

import java.util.Arrays;
import java.util.List;


public class myclass {
    /** Application name. */
    private static final String APPLICATION_NAME =
            "CS 441 HW1";

    /** Directory to store user credentials for this application. */
    private static final java.io.File DATA_STORE_DIR = new java.io.File(
            System.getProperty("user.home"), ".credentials/hw1");

    /** Global instance of the {@link FileDataStoreFactory}. */
    private static FileDataStoreFactory DATA_STORE_FACTORY;

    /** Global instance of the JSON factory. */
    private static final JsonFactory JSON_FACTORY =
            JacksonFactory.getDefaultInstance();

    /** Global instance of the HTTP transport. */
    private static HttpTransport HTTP_TRANSPORT;

    /** Global instance of the scopes required by this quickstart.
     *
     * If modifying these scopes, delete your previously saved credentials
     * at ~/.credentials/gmail-java-quickstart
     */
    private static final List<String> SCOPES =
            Arrays.asList(GmailScopes.GMAIL_READONLY,GmailScopes.GMAIL_MODIFY , SheetsScopes.SPREADSHEETS);

    //Scopes required for this project

    static {
        try {
            HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
            DATA_STORE_FACTORY = new FileDataStoreFactory(DATA_STORE_DIR);
        } catch (Throwable t) {
            t.printStackTrace();
            System.exit(1);
        }
    }

    /**
     * Creates an authorized Credential object.
     * @return an authorized Credential object.
     * @throws IOException
     */
    public static Credential authorize() throws IOException {
        // Load client secrets.
        InputStream in =
                myclass.class.getResourceAsStream("/client_secret.json");
        GoogleClientSecrets clientSecrets =
                GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

        // Build flow and trigger user authorization request.
        GoogleAuthorizationCodeFlow flow =
                new GoogleAuthorizationCodeFlow.Builder(
                        HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES)
                        .setDataStoreFactory(DATA_STORE_FACTORY)
                        .setAccessType("offline")
                        .build();
        Credential credential = new AuthorizationCodeInstalledApp(
                flow, new LocalServerReceiver()).authorize("user");
        System.out.println(
                "Credentials saved to " + DATA_STORE_DIR.getAbsolutePath());
        return credential;
    }

    /**
     * Build and return an authorized Gmail client service.
     * @return an authorized Gmail client service
     * @throws IOException
     */
    public static Gmail getGmailService() throws IOException {
        Credential credential = authorize();
        return new Gmail.Builder(HTTP_TRANSPORT, JSON_FACTORY, credential)
                .setApplicationName(APPLICATION_NAME)
                .build();
    }

    public static Sheets getSheetsService() throws IOException {
        Credential credential = authorize();
        return new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, credential)
                .setApplicationName(APPLICATION_NAME)
                .build();
    }


    //Main method for the class

    public static void main(String[] args) throws IOException, ParseException {

        //Setting up services for gmail and sheets API
        Gmail gmailservice = getGmailService();
        Sheets sheetsservice = getSheetsService();

        String user = "me";

        //Getting all unread messages in the inbox of the gmail account

        ListMessagesResponse response = gmailservice.users().messages().list("me").setQ("in:inbox is:unread").execute();

                List<Message> messages = response.getMessages();

        //Setting parameters for getting the responses that are stored later
        List<String> abc=new ArrayList();
        abc.add("subject");
        abc.add("received");



        //Date formats to parse the data obtained from the message header
        SimpleDateFormat ft = new SimpleDateFormat ("dd MMM yyyy hh:mm:ss Z (z)");
        SimpleDateFormat dateoutput = new SimpleDateFormat ("ddMMyyyy");
        SimpleDateFormat timeoutput = new SimpleDateFormat("HH:mm:ss");
        //date formats for getting the date and time to store in the sheets

        Date t;
        String sheet_id="";
        List<String> removeLabel=new ArrayList();

        //marking messages received as read
        removeLabel.add("UNREAD");

        String date="";
        String previous_date="";
        int i=0;

        //only enter below code if there are unread messages
        if(!(messages==null)) {
            for (Message message : messages) {

                //each message is marked as read
                ModifyThreadRequest mods = new ModifyThreadRequest()
                        .setRemoveLabelIds(removeLabel);

                //message metadata is stored in the object below
                Message message1 = gmailservice.users().messages().get(user, message.getId()).setFormat("metadata").setMetadataHeaders(abc).execute();

                //date is retrieved from the JSON response
                 date = message1.getPayload().getHeaders().get(0).getValue().split(";")[1].split(",")[1];

                t = ft.parse(date.trim());

                List<Request> requests = new ArrayList<>();
                List<CellData> values = new ArrayList<>();


                //if the message's date is a new date then a new sheet is created
                if(!previous_date.equals(dateoutput.format(t)))

                {
                    i=0;
                    previous_date = dateoutput.format(t);

                    //sheet title is set to the date of the message
                    Spreadsheet sheet = new Spreadsheet();
                    sheet.setProperties(new SpreadsheetProperties().setTitle(previous_date));

                    Spreadsheet created_sheet = sheetsservice.spreadsheets().create(sheet).execute();
                    sheet_id = created_sheet.getSpreadsheetId();

        //each sheet is populated with "time" and "subject" as headings
                    values.add(new CellData()
                            .setUserEnteredValue(new ExtendedValue().setStringValue("TIME")));
                    values.add(new CellData()
                            .setUserEnteredValue(new ExtendedValue().setStringValue("SUBJECT")));


                    //above value are added to a request which is executed later
                    requests.add(new Request()
                            .setUpdateCells(new UpdateCellsRequest()
                                    .setStart(new GridCoordinate()
                                            .setSheetId(0)
                                            .setRowIndex(0)
                                            .setColumnIndex(i))
                                    .setRows(Arrays.asList(
                                            new RowData().setValues(values)))
                                    .setFields("userEnteredValue")));
                    i++;


                    BatchUpdateSpreadsheetRequest batchUpdateRequest = new BatchUpdateSpreadsheetRequest()
                            .setRequests(requests);
                    sheetsservice.spreadsheets().batchUpdate(sheet_id, batchUpdateRequest)
                            .execute();
                }
                values.clear();
                requests.clear();


                // new data(time and subject) is written to a new row specified by i
                values.add(new CellData()
                            .setUserEnteredValue(new ExtendedValue().setStringValue(timeoutput.format(t))));

                    values.add(new CellData()
                            .setUserEnteredValue(new ExtendedValue().setStringValue(message1.getPayload().getHeaders().get(message1.getPayload().getHeaders().size() - 1).getValue())));

                requests.add(new Request()
                            .setUpdateCells(new UpdateCellsRequest()
                                    .setStart(new GridCoordinate()
                                            .setSheetId(0)
                                            .setRowIndex(i)
                                            .setColumnIndex(0))
                                    .setRows(Arrays.asList(
                                             new RowData().setValues(values)))
                                    .setFields("userEnteredValue")));




                i++;

            //the request is executed below
                BatchUpdateSpreadsheetRequest batchUpdateRequest = new BatchUpdateSpreadsheetRequest()
                        .setRequests(requests);
                sheetsservice.spreadsheets().batchUpdate(sheet_id, batchUpdateRequest)
                        .execute();
            }
        }

    }


}