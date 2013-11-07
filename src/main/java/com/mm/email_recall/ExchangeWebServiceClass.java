package com.mm.email_recall;

import java.io.FileInputStream;
import java.net.URI;
import java.util.Properties;

import microsoft.exchange.webservices.data.Appointment;
import microsoft.exchange.webservices.data.ConnectingIdType;
import microsoft.exchange.webservices.data.EmailMessage;
import microsoft.exchange.webservices.data.ExchangeService;
import microsoft.exchange.webservices.data.ImpersonatedUserId;
import microsoft.exchange.webservices.data.Item;
import microsoft.exchange.webservices.data.WebCredentials;
import microsoft.exchange.webservices.data.ItemView;
import microsoft.exchange.webservices.data.WellKnownFolderName;
import microsoft.exchange.webservices.data.FindItemsResults;

public class ExchangeWebServiceClass {
	public static void main(String[] args) {		
		Properties prop = new Properties();
		try {
			prop.load(new FileInputStream("Credentials.properties"));
		    String Username = prop.getProperty("username");
		    String Password = prop.getProperty("password");
		    String ExchangeDomain = prop.getProperty("domain");
		    String ExchangeURI = prop.getProperty("exchangeuri");
			ExchangeService service = new ExchangeService();
//			URI url = new URI("https://24.185.27.191/EWS/exchange.asmx");
			URI url = new URI(ExchangeURI);
			service.setUrl(url);					
			service.setCredentials( new WebCredentials(Username, Password, ExchangeDomain ) );  //this works			
			System.out.println( "Created ExchangeService" );
//			service.setTraceEnabled(true);
			String[] names = {"alan.reid@makesoftlab.local","administrator@makesoftlab.local","aisha.bhari@makesoftlab.local","alannah.shaw@makesoftlab.local"};

			for (String sname: names)
			{
				ImpersonatedUserId impersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, sname);
				service.setImpersonatedUserId(impersonatedUserId); 				
				ItemView view = new ItemView(10000);
				
// Loop through Inbox
				FindItemsResults<Item> findResults = service.findItems(WellKnownFolderName.Inbox, view);
				
				for(Item item : findResults.getItems())
				{
					// Bind to an existing message using its unique identifier.
					EmailMessage message = EmailMessage.bind(service, item.getId());
	//				item.load();
					String[] allrecipients = item.getDisplayTo().split(";");
					for (String tmprecipients: allrecipients)
						{
//							System.out.println("ID:" + item.getId() + "~Mailbox:" + sname + "~Sender:" + message.getSender().getName() + "~To:" + tmprecipients + "~Subject:" + item.getSubject() + "~Date sent:" + item.getDateTimeReceived());
						System.out.println("Inbox:" + sname + "~Sender:" + message.getSender().getName() + "~To:" + tmprecipients + "~Subject:" + item.getSubject() + "~Date sent:" + item.getDateTimeReceived());
						}
					}
				// Loop through Calendar
				findResults = service.findItems(WellKnownFolderName.SentItems, view);	
				for(Item item : findResults.getItems())
				{
					// Bind to an existing message using its unique identifier.
					EmailMessage message = EmailMessage.bind(service, item.getId());
	//				item.load();
					String[] allrecipients = item.getDisplayTo().split(";");
					for (String tmprecipients: allrecipients)
						{
//							System.out.println("ID:" + item.getId() + "~Mailbox:" + sname + "~Sender:" + message.getSender().getName() + "~To:" + tmprecipients + "~Subject:" + item.getSubject() + "~Date sent:" + item.getDateTimeReceived());
						System.out.println("Sent:" + sname + "~Sender:" + message.getSender().getName() + "~To:" + tmprecipients + "~Subject:" + item.getSubject() + "~Date sent:" + item.getDateTimeReceived());
						}
				}					

// Loop through Calendar
				findResults = service.findItems(WellKnownFolderName.Calendar, view);	
				for(Item itemCal : findResults.getItems())
				{					
					Appointment appointment = (Appointment) itemCal;
					String[] allattendees = itemCal.getDisplayTo().split(";");
					for (String tmpattendees: allattendees)
						{
						
						System.out.println("Calendar:" + sname + "~Organizer:" + appointment.getOrganizer().getName() + "~Attendees:" + tmpattendees + "~Subject:" + appointment.getSubject() + "~Start:" + appointment.getStart() + "~End:" + appointment.getEnd());
						}		
				}					
			} 
		}
		catch (Exception e) {			
		e.printStackTrace();
		}
	}
}
