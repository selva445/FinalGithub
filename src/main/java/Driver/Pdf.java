package Driver;

import java.io.BufferedInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.util.Date;
import java.util.HashMap;
import java.util.Hashtable;
import java.util.Map;
import java.util.Properties;
import java.util.concurrent.TimeUnit;

import javax.mail.Flags;
import javax.mail.Folder;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Session;
import javax.mail.Store;
import javax.mail.Flags.Flag;
import javax.mail.search.FlagTerm;

import org.apache.commons.mail.DefaultAuthenticator;
import org.apache.commons.mail.EmailAttachment;
import org.apache.commons.mail.MultiPartEmail;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.safari.SafariDriver;
import org.sikuli.script.FindFailed;
import org.sikuli.script.Pattern;
import org.sikuli.script.Screen;

import utility.Constant;
import utility.ExcelUtils;
import utility.Extend_Report;
import utility.Log;

public class Pdf {

	public static void main(String[] args) throws Exception {
	/*	// TODO Auto-generated method stub
		// Creating Object of 'Screen' class
		 //Screen is a base class provided by Sikuli. It allows us to access all the methods provided by Sikuli.
		 Screen screen = new Screen();
		 // Creating Object of Pattern class and specify the path of specified images
		 // I have captured images of Facebook Email id field, Password field and Login button and placed in my local directory
		 // Facebook user id image 
		 Pattern username = new Pattern("C:\\Users\\889145\\Desktop\\Front_Car.PNG");
		 // Facebook password image
		 
		 // Facebook login button image
		 
		 // Initialization of driver object to launch firefox browser 
		// System.setProperty("webdriver.gecko.driver", System.getProperty("user.dir")+"\\src\\drivers\\geckodriver.exe");
		// WebDriver driver = new FirefoxDriver();
		 System.setProperty("webdriver.ie.driver",Constant.Path_IEDriver);
			DesiredCapabilities capabilities = DesiredCapabilities.internetExplorer();
			capabilities.setCapability(InternetExplorerDriver.INITIAL_BROWSER_URL,Constant.SYS_URL);							//
			capabilities.setCapability(InternetExplorerDriver.INTRODUCE_FLAKINESS_BY_IGNORING_SECURITY_DOMAINS, true);
			capabilities.setCapability(InternetExplorerDriver.IE_ENSURE_CLEAN_SESSION, true);
			capabilities.setCapability("requireWindowFocus", true);
			//capabilities.setCapability("ignoreZoomSetting", true); 
			//capabilities.setCapability("nativeEvents", false);    
			capabilities.setCapability("unexpectedAlertBehaviour", "accept");
			capabilities.setCapability("ignoreProtectedModeSettings", true);
			capabilities.setCapability("disable-popup-blocking", true);
			capabilities.setCapability("enablePersistentHover", true);
		//	WebDriver driver = new FirefoxDriver();
			WebDriver driver = new InternetExplorerDriver(capabilities);
		 System.setProperty("webdriver.ie.driver",Constant.Path_IEDriver);
		
		 // To maximize the browser
		 driver.manage().window().maximize();
		 // Open Facebook
		 driver.get("https://audanet-cae.audatex.com.au/sso/login?service=https%3a%2f%2faudanet-cae.audatex.com.au%2fbre%2fwork%2fNO_PROCESS%2fDashboard%3fpageViewpoint%3dMAIN");
		
		 driver.findElement(By.id("ssousername")).sendKeys("qbeSelvakumar");
		 driver.findElement(By.id("password")).sendKeys("EwUuQhp4uC");
		 driver.findElement(By.name("submit")).click();
		 driver.findElement(By.xpath("(.//*[normalize-space(text()) and normalize-space(.)='Welcome back to AudaNet'])[1]/following::i[1]")).click();
		 
		 Thread.sleep(5000);
		 WebElement Edit_Username= driver.findElement(By.xpath("(.//*[normalize-space(text()) and normalize-space(.)='Welcome back to AudaNet'])[1]/following::i[1]"));
		 Actions action = new Actions(driver);
		 //Find the targeted element
		 
		                //Here I used JavascriptExecutor interface to scroll down to the targeted element
		 ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView();", Edit_Username);
		                //used doubleClick(element) method to do double click action
		 action.doubleClick(Edit_Username).build().perform();
		 
		 Thread.sleep(2000);
		 driver.findElement(By.id("toDoListItem_QBEVehicledata")).click();
		 Thread.sleep(2000);
		 driver.findElement(By.id("BREForm_root.task.basicClaimData.vehicle.advancedVehicleDamage.claimDamageMode")).click();
		 Thread.sleep(2000);
		 // Calling 'type' method to enter username in the email field using 'screen' object
		 //screen.click(username);

		 String sikulipart = Constant.Sikuli_Path;	
		 
		 
		String Carclick = sikulipart.concat("BackCar"+".PNG");
		Pattern Carclickp = new Pattern(Carclick);	
		screen.find(Carclickp);
		 screen.click(Carclickp);*/
		
		/* String val1="Repairer Allocated to Request - Test Garage 2 PH2";
		              
		 
		String sUserName1="Repairer Allocated to Request - Test Garage 2 PH2";
		boolean val2 = val1.contains(sUserName1);
		
		
		System.out.println(val2);*/
		
	        
	        // page with example pdf document
	      //  driver.get("file:///C:/Users/selvakumar.c/Downloads/Fram6/Fram6 - Copy/src/main/java/Upload/Quote 1.pdf");
	        URL url = new URL("file:///C:/Users/selvakumar.c/Downloads/Fram6/Fram6 - Copy/src/main/java/Upload/Quote 1.pdf");
	        InputStream is = url.openStream();
	        BufferedInputStream fileToParse = new BufferedInputStream(is);
	        PDDocument document = null;
	        try {
	            document = PDDocument.load(fileToParse);
	            String output = new PDFTextStripper().getText(document);
	            System.out.println(output);
	        } finally {
	            if (document != null) {
	                document.close();
	            }
	            fileToParse.close();
	            is.close();
	        }
	       
	    }
		    
	
	/*	Store store = null;
		MultiPartEmail et = new MultiPartEmail();
		 Properties props = new Properties();
	     props.put("mail.store.protocol", "imaps");
	     Session session = Session.getInstance(props);
	     store = session.getStore("imaps");
	     store.connect("imap.outlook.com", Constant.Outlook_EmailID, Constant.Outlook_EmailPwd);
	     et.setHostName("smtp-mail.outlook.com");
       et.setSmtpPort(465);
       et.setAuthenticator(new DefaultAuthenticator(Constant.Outlook_EmailID, Constant.Outlook_EmailPwd));
        et.setSSLOnConnect(true);
        et.setFrom("Selvakumar.c@mphasis.com");
        et.setSubject("TestMail");
        et.setMsg("This is a test mail ... :-)");
        et.addTo("foo@bar.com");
        EmailAttachment attachment = new EmailAttachment();
        attachment.setPath("C:\\Users\\889145\\Desktop\\LMIwork\\Fram6\\src\\main\\java\\testData\\Run_Manager.xlsx");
        attachment.setDisposition(EmailAttachment.ATTACHMENT);
        attachment.setDescription("Picture of Result");
        attachment.setName("Run_Manager");
        et.attach(attachment);
        et.send();*/
        
       /* Application oApp = new Application();

        //Create the new message by using the simplest approach.
        MailItem oMsg = (MailItem)oApp.CreateItem(OlItemType.olMailItem);

        //Add a recipient.
        // TODO: Change the following recipient where appropriate.
        Recipient oRecip = (Recipient)oMsg.Recipients.Add(recepientAddress);
        oRecip.Resolve();

        //Set the basic properties.
        oMsg.Subject = messageSubject;
        oMsg.Body = messageBody;

        //Add an attachment.
        // TODO: change file path where appropriate
        String sSource = attachmentfile;
        String sDisplayName = attachmentDisplayName;
        int iPosition = (int)oMsg.Body.Length + 1;
        int iAttachType = (int)OlAttachmentType.olByValue;  
        Attachment oAttach = oMsg.Attachments.Add(sSource,iAttachType,iPosition,sDisplayName);

        // If you want to, display the message.
        //oMsg.Display(true);  //modal

        //Send the message.
        oMsg.Save();
        oMsg.Send();
      

        //Explicitly release objects.
        oRecip = null;
        oAttach = null;
        oMsg = null;
        oApp = null;
		
		*/
	}


