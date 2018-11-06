package utility;

import org.openqa.selenium.WebDriver;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

public class Extend_Report {
	

	public static String path;
		static ExtentReports extent;
		static ExtentTest logger;
	
	
	public  static void Set_Header(String  TCname) throws Exception{
		
		try{
			String Path=System.getProperty("user.dir") +"//src//main//java//Html_Reports/";		
			
			String TC = Path+TCname+"Test.html";
			
			extent = new ExtentReports (TC, true);		
			logger = extent.startTest(TCname);	
		}
		catch(Exception e)
		{
			System.out.println("Exception In Setting Header"+ e.toString());
		}
		
	}
	
	
	
	public  static void AddReport(String  Text, String Status) throws Exception{
		try
		{
			if(Status.equalsIgnoreCase("Pass"))			
			{			
						
				
				 logger.log(LogStatus.PASS, Text);
				
			}
			else if (Status.equalsIgnoreCase("Fail"))
			
			{
				
				logger.log(LogStatus.FAIL, Text);
				
				
			}
			
			else 
				
			{
				
				
				logger.log(LogStatus.INFO, Text);	
			}
		}
		catch(Exception e)
		{
			System.out.println(e.toString());
			
			
		}
		
		
		
		
	}
	
	public  static void Clean_Up() throws Exception{
		
		
		try
		{	
				
			extent.endTest(logger);
			
			extent.flush();
			
			extent.close();
	      
		}
		catch(Exception e)
		{
			System.out.println("Exception"+ e.toString());
		}
	}
	
}
