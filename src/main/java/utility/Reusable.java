package utility;

import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import java.sql.Driver;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;
import java.util.concurrent.TimeUnit;
//import java.util.regex.Pattern;































import javax.annotation.Resource;
import javax.imageio.ImageIO;
import javax.mail.Flags;
import javax.mail.Flags.Flag;
import javax.mail.Folder;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Session;
import javax.mail.Store;
import javax.mail.search.FlagTerm;

import org.apache.commons.io.FileUtils;
import org.apache.commons.mail.DefaultAuthenticator;
import org.apache.commons.mail.Email;
import org.apache.commons.mail.EmailAttachment;
import org.apache.commons.mail.MultiPartEmail;
import org.apache.commons.mail.SimpleEmail;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
//import org.apache.poi.sl.draw.geom.Path;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.Point;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxProfile;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.sikuli.script.FindFailed;
import org.sikuli.script.Pattern;
import org.sikuli.script.Screen;

import pageObjects.BaseClass;
import Repository.RepositoryParser;

public class Reusable {
	
	//******************************************************************************************************************************//
	public static void Enter_Value_WebElement(WebDriver driver,int iTestCaseRow,String object_value, String Excel_Workbook,String TC_Name) throws Exception
	{
	
			
		
		try
		{
			
	    	int val_Uname=ExcelUtils.get_column(object_value,Excel_Workbook);
	    	
	    	String sUserName1 = ExcelUtils.getCellData(iTestCaseRow, val_Uname);  
	    	
	    	String sUserName = sUserName1.trim();
	    	
	    	
	    	
	    	if(sUserName != null && !sUserName.isEmpty()) 
	    	{
	    		Thread.sleep(1000);
	    		WebElement Edit_Username=RepositoryParser.Create_Objects(object_value,Excel_Workbook); 	
	    		JavascriptExecutor js = (JavascriptExecutor) driver;
	    		js.executeScript("var evt = document.createEvent('MouseEvents');" + "evt.initMouseEvent('click',true, true, window, 0, 0, 0, 0, 0, false, false, false, false, 0,null);" + "arguments[0].dispatchEvent(evt);", Edit_Username);
	    		//Edit_Username.click();
		    	Edit_Username.clear();	    		
	    		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
	    		
	    		String Label = Edit_Username.getAttribute("id");
	    		Edit_Username.sendKeys(sUserName);
	    		Edit_Username.sendKeys(Keys.TAB);
	    		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
	            Log.info(sUserName+" is entered in text box" );	           
	            Extend_Report.AddReport( sUserName +" is Entered in the "+ Label +" field", "Pass");
	            Reusable.takeScreenshot(driver, object_value,TC_Name);
	           
	    	}
		}
    	
    	catch(Exception e)
		{	
    		Reusable.takeScreenshot(driver, object_value,TC_Name);
    		ExcelUtils.setExcelFile(Constant.Path_Excel+Constant.File_RunManager, "Run_Manager");		 
    		int val_run=ExcelUtils.getRowContains_manager(TC_Name, 0, "Run_Manager");		 
    		ExcelUtils.setCellData_Manager("FAIL", val_run, 2);
    		Constant.TC_Status= "FAIL";    		
    		 Extend_Report.AddReport("Error in Enter_Value_WebElement", "Fail");
    		System.out.println("Error in Enter_Value_WebElement " + e.toString());
    		Log.info(" Error in entered in text box" );
		}
		
	}
	
	//******************************************************************************************************************************//
		public static void Clear_Value_WebElement(WebDriver driver,int iTestCaseRow,String object_value, String Excel_Workbook,String TC_Name) throws Exception
		{
		
				
			
			try
			{
				
		    	int val_Uname=ExcelUtils.get_column(object_value,Excel_Workbook);
		    	
		    	String sUserName1 = ExcelUtils.getCellData(iTestCaseRow, val_Uname);  
		    	
		    	String sUserName = sUserName1.trim();
		    	
		    	
		    	
		    	if(sUserName != null && !sUserName.isEmpty()) 
		    	{
		    		Thread.sleep(1000);
		    		WebElement Edit_Username=RepositoryParser.Create_Objects(object_value,Excel_Workbook); 	
		    		JavascriptExecutor js = (JavascriptExecutor) driver;
		    		js.executeScript("var evt = document.createEvent('MouseEvents');" + "evt.initMouseEvent('click',true, true, window, 0, 0, 0, 0, 0, false, false, false, false, 0,null);" + "arguments[0].dispatchEvent(evt);", Edit_Username);
		    		//Edit_Username.click();
			    	Edit_Username.clear();	    		
		    		
		    		Edit_Username.sendKeys(Keys.TAB);
		    		
		            Log.info(sUserName+" is entered in text box" );	  
		            
		            Extend_Report.AddReport( "Field :" +Edit_Username.toString() +" is Cleared ", "Pass");
		            Reusable.takeScreenshot(driver, object_value,TC_Name);
		           
		    	}
			}
	    	
	    	catch(Exception e)
			{	
	    		Reusable.takeScreenshot(driver, object_value,TC_Name);
	    		ExcelUtils.setExcelFile(Constant.Path_Excel+Constant.File_RunManager, "Run_Manager");		 
	    		int val_run=ExcelUtils.getRowContains_manager(TC_Name, 0, "Run_Manager");		 
	    		ExcelUtils.setCellData_Manager("FAIL", val_run, 2);
	    		Constant.TC_Status= "FAIL";    
	    		 Extend_Report.AddReport("Error in Enter_Value_WebElement", "Fail");
	    		System.out.println("Error in Enter_Value_WebElement " + e.toString());
	    		Log.info(" Error in entered in text box" );
			}
			
		}
	//******************************************************************************************************************************//
		public static void Enter_Value_WebElementConstant(WebDriver driver,int iTestCaseRow,String object_value, String Excel_Workbook,String TC_Name) throws Exception
		{
		
				
			
			try
			{
				
		    	int val_Uname=ExcelUtils.get_column(object_value,Excel_Workbook);
		    	
		    	String sUserName1 = ExcelUtils.getCellData(iTestCaseRow, val_Uname);  
		    	
		    	String sUserName = sUserName1.trim();
		    	
		    	
		    	
		    	if(sUserName != null && !sUserName.isEmpty()) 
		    	{
		    		Thread.sleep(1000);
		    		WebElement Edit_Username=RepositoryParser.Create_Objects(object_value,Excel_Workbook); 	
		    		JavascriptExecutor js = (JavascriptExecutor) driver;
		    		js.executeScript("var evt = document.createEvent('MouseEvents');" + "evt.initMouseEvent('click',true, true, window, 0, 0, 0, 0, 0, false, false, false, false, 0,null);" + "arguments[0].dispatchEvent(evt);", Edit_Username);
		    		//Edit_Username.click();
			    	Edit_Username.clear();	    		
		    		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		    		
		    		String Label = Edit_Username.getAttribute("id");
		    		sUserName=Constant.Global_ClaimNO;
		    		Edit_Username.sendKeys(sUserName);
		    		Edit_Username.sendKeys(Keys.TAB);
		    		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		            Log.info(sUserName+" is entered in text box" );	           
		            Extend_Report.AddReport( sUserName +" is Entered in the "+ Label +" field", "Pass");
		            Reusable.takeScreenshot(driver, object_value,TC_Name);
		           
		    	}
			}
	    	
	    	catch(Exception e)
			{	
	    		Reusable.takeScreenshot(driver, object_value,TC_Name);
	    		ExcelUtils.setExcelFile(Constant.Path_Excel+Constant.File_RunManager, "Run_Manager");		 
	    		int val_run=ExcelUtils.getRowContains_manager(TC_Name, 0, "Run_Manager");		 
	    		ExcelUtils.setCellData_Manager("FAIL", val_run, 2);
	    		Constant.TC_Status= "FAIL";    
	    		 Extend_Report.AddReport("Error in Enter_Value_WebElement", "Fail");
	    		System.out.println("Error in Enter_Value_WebElement " + e.toString());
	    		Log.info(" Error in entered in text box" );
			}
			
		}
	//******************************************************************************************************************************//
		public static void Enter_Value_WebElement_Set(WebDriver driver,int iTestCaseRow,String object_value, String Excel_Workbook,String TC_Name) throws Exception
		{
		
				
			
			try
			{
				
		    	int val_Uname=ExcelUtils.get_column(object_value,Excel_Workbook);
		    	
		    	String sUserName1 = ExcelUtils.getCellData(iTestCaseRow, val_Uname);  
		    	
		    	String sUserName = sUserName1.trim();
		    	
		    	
		    	
		    	if(sUserName != null && !sUserName.isEmpty()) 
		    	{
		    		Thread.sleep(1000);
		    		WebElement Edit_Username=RepositoryParser.Create_Objects(object_value,Excel_Workbook); 	
		    		//JavascriptExecutor js = (JavascriptExecutor) driver;
		    	//	js.executeScript("var evt = document.createEvent('MouseEvents');" + "evt.initMouseEvent('click',true, true, window, 0, 0, 0, 0, 0, false, false, false, false, 0,null);" + "arguments[0].dispatchEvent(evt);", Edit_Username);
		    		//Edit_Username.click();
			    	Edit_Username.clear();	    		
		    		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		    		
		    		String Label = Edit_Username.getAttribute("id");
		    		Edit_Username.sendKeys(sUserName);
		    		Edit_Username.sendKeys(Keys.TAB);
		    		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		            Log.info(sUserName+" is entered in text box" );	           
		            Extend_Report.AddReport( sUserName +" is Entered in the "+ Label +" field", "Pass");
		            Reusable.takeScreenshot(driver, object_value,TC_Name);
		            
		    	}
			}
	    	
	    	catch(Exception e)
			{	
	    		Reusable.takeScreenshot(driver, object_value,TC_Name);
	    		ExcelUtils.setExcelFile(Constant.Path_Excel+Constant.File_RunManager, "Run_Manager");		 
	    		int val_run=ExcelUtils.getRowContains_manager(TC_Name, 0, "Run_Manager");		 
	    		ExcelUtils.setCellData_Manager("FAIL", val_run, 2);
	    		Constant.TC_Status= "FAIL";    
	    		 Extend_Report.AddReport("Error in Enter_Value_WebElement", "Fail");
	    		System.out.println("Error in Enter_Value_WebElement " + e.toString());
	    		Log.info(" Error in entered in text box" );
			}
			
		}
	
	//******************************************************************************************************************************//
		public static void Enter_Value_WebElement_wait(WebDriver driver,int iTestCaseRow,String object_value, String Excel_Workbook,String TC_Name) throws Exception
		{
		
				
			
			try
			{
				
		    	int val_Uname=ExcelUtils.get_column(object_value,Excel_Workbook);
		    	
		    	String sUserName1 = ExcelUtils.getCellData(iTestCaseRow, val_Uname);  
		    	
		    	String sUserName = sUserName1.trim();
		    
		    	if(sUserName != null && !sUserName.isEmpty()) 
		    	{
		    		Thread.sleep(10000);
		    		WebElement Edit_Username=RepositoryParser.Create_Objects(object_value,Excel_Workbook); 	
		    		JavascriptExecutor js = (JavascriptExecutor) driver;
		    		js.executeScript("var evt = document.createEvent('MouseEvents');" + "evt.initMouseEvent('click',true, true, window, 0, 0, 0, 0, 0, false, false, false, false, 0,null);" + "arguments[0].dispatchEvent(evt);", Edit_Username);
		    		//Edit_Username.click();
			    	Edit_Username.clear();	    		
		    		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		    		
		    		String Label = Edit_Username.getAttribute("name");
		    		Edit_Username.sendKeys(sUserName);
		    		Edit_Username.sendKeys(Keys.TAB);
		    		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		            Log.info(sUserName+" is entered in text box" );
		            Extend_Report.AddReport( sUserName +" is Entered in the "+ Label +" field", "Pass");
		            Reusable.takeScreenshot(driver, object_value,TC_Name);
		            
		    	}
			}
	    	
	    	catch(Exception e)
			{	
	    		Reusable.takeScreenshot(driver, object_value,TC_Name);
	    		ExcelUtils.setExcelFile(Constant.Path_Excel+Constant.File_RunManager, "Run_Manager");		 
	    		int val_run=ExcelUtils.getRowContains_manager(TC_Name, 0, "Run_Manager");		 
	    		ExcelUtils.setCellData_Manager("FAIL", val_run, 2);
	    		Constant.TC_Status= "FAIL";    
	    		 Extend_Report.AddReport("Error is  Enter_Value_WebElement", "Fail");
	    		System.out.println("Error in Enter_Value_WebElement " + e.toString());
	    		Log.info(" Error in entered in text box" );
			}
			
		}
	//******************************************************************************************************************************//
		public static void MouseOver_WebElement(WebDriver driver,int iTestCaseRow,String object_value, String Excel_Workbook,String TC_Name) throws Exception
		{
		
				
			
			try
			{
				
		    	         	
		    	int val_Uname=ExcelUtils.get_column(object_value,Excel_Workbook);   
		    	
		    	String sUserName1 = ExcelUtils.getCellData(iTestCaseRow, val_Uname);  
		    	String sUserName = sUserName1.trim();
		    
		    	
		    	if(sUserName != null && !sUserName.isEmpty()) 
		    	{
		    		Thread.sleep(2000);
		    		WebElement Edit_Username=RepositoryParser.Create_Objects(object_value,Excel_Workbook); 	
		    		Actions action = new Actions(driver);
		    		System.out.println(Edit_Username.toString());
		    		action.moveToElement(Edit_Username).build().perform();
		            Log.info(Edit_Username+"MouseOver Operation is done over"+Edit_Username );
		            String Val = Edit_Username.getText();
		            if (Val=="")
		            {
		            	Extend_Report.AddReport( "MouseOver Operation is done over"+Edit_Username , "Pass");
		            }
		            else
		            {
		            	Extend_Report.AddReport( "MouseOver Operation is done over"+Edit_Username.getText() , "Pass");
		            }
		            
		            Reusable.takeScreenshot(driver, object_value,TC_Name);
		            
		    	}
			}
	    	
	    	catch(Exception e)
			{	
	    		Reusable.takeScreenshot(driver, object_value,TC_Name);
	    		ExcelUtils.setExcelFile(Constant.Path_Excel+Constant.File_RunManager, "Run_Manager");		 
	    		int val_run=ExcelUtils.getRowContains_manager(TC_Name, 0, "Run_Manager");		 
	    		ExcelUtils.setCellData_Manager("FAIL", val_run, 2);
	    		Constant.TC_Status= "FAIL";    
	    		Extend_Report.AddReport("Error is  MouseOver_WebElement", "Fail");
	    		System.out.println("Error in MouseOver_WebElement " + e.toString());
	    		
			}
			
		}
		//******************************************************************************************************************************//
				public static void MouseOver_WebElementClick(WebDriver driver,int iTestCaseRow,String object_value, String Excel_Workbook,String TC_Name) throws Exception
				{
				
						
					
					try
					{
						
				    	         	
				    	int val_Uname=ExcelUtils.get_column(object_value,Excel_Workbook); 				    	
				    	String sUserName1 = ExcelUtils.getCellData(iTestCaseRow, val_Uname);  
				    	String sUserName = sUserName1.trim();
				    	
				    	
				    	if(sUserName != null && !sUserName.isEmpty()) 
				    	{
				    		Thread.sleep(2000);
				    		WebElement Edit_Username=RepositoryParser.Create_Objects(object_value,Excel_Workbook); 				    		
				    		
				    		JavascriptExecutor js = (JavascriptExecutor) driver;
				    		js.executeScript("var evt = document.createEvent('MouseEvents');" + "evt.initMouseEvent('click',true, true, window, 0, 0, 0, 0, 0, false, false, false, false, 0,null);" + "arguments[0].dispatchEvent(evt);", Edit_Username);			    		
				            Log.info(Edit_Username+"MouseOver Operation and click is done over"+Edit_Username );
				            Extend_Report.AddReport( "MouseOver Operation and click is done over"+Edit_Username.toString() , "Pass");
				            Reusable.takeScreenshot(driver, object_value,TC_Name);
				            
				    	}
					}
			    	
			    	catch(Exception e)
					{	
			    		Reusable.takeScreenshot(driver, object_value,TC_Name);
			    		ExcelUtils.setExcelFile(Constant.Path_Excel+Constant.File_RunManager, "Run_Manager");		 
			    		int val_run=ExcelUtils.getRowContains_manager(TC_Name, 0, "Run_Manager");		 
			    		ExcelUtils.setCellData_Manager("FAIL", val_run, 2);
			    		Constant.TC_Status= "FAIL";    
			    		Extend_Report.AddReport("Error is  MouseOver_WebElement", "Fail");
			    		System.out.println("Error in MouseOver_WebElement " + e.toString());
			    		
					}
					
				}
				
				//******************************************************************************************************************************//
				public static void ExportExcel(WebDriver driver,int iTestCaseRow,String object_value, String Excel_Workbook,String TC_Name) throws Exception
				{
				
						
					
					try
					{
						
				    	         	
				    	int val_Uname=ExcelUtils.get_column(object_value,Excel_Workbook); 				    	
				    	String sUserName1 = ExcelUtils.getCellData(iTestCaseRow, val_Uname);  
				    	String sUserName = sUserName1.trim();
				    	
				    	
				    	if(sUserName != null && !sUserName.isEmpty()) 
				    	{
				    		if (sUserName.contains("QBE"))
				    		{
				    			
				    			Thread.sleep(2000);
					    		WebElement Edit_Username=RepositoryParser.Create_Objects(object_value,Excel_Workbook); 				    		
					    		String val=Edit_Username.getText();
					    		System.out.println(val);
					    		String val1=val.substring(52, 63);
					    		System.out.println(val1);
					    		Constant.Global_ClaimNO=val1;
					    		ExcelUtils.setCellData(val1, iTestCaseRow, val_Uname);		    		
					            Log.info(Edit_Username+"ExportExcel Done For "+Edit_Username );
					            Extend_Report.AddReport( "ExportExcel Done For  "+Edit_Username.toString() , "Pass");
					            Reusable.takeScreenshot(driver, object_value,TC_Name);
				    			
				    		}
				    		else
				    		{
				    			
				    		
				    		Thread.sleep(2000);
				    		WebElement Edit_Username=RepositoryParser.Create_Objects(object_value,Excel_Workbook); 				    		
				    		String val=Edit_Username.getText();
				    		if (val.length()>60) 
				    		{
				    			Select Edit_Username2=RepositoryParser.Create_Objects_Select(object_value,Excel_Workbook);
				    			System.out.println( Edit_Username2.toString());
				    			WebElement option1 = Edit_Username2.getFirstSelectedOption();
				    			String defaultItem = option1.getText();
				    			ExcelUtils.setCellData(defaultItem, iTestCaseRow, val_Uname);		    		
					            Log.info(Edit_Username+"ExportExcel Done For "+defaultItem );
					            Extend_Report.AddReport( "ExportExcel Done For  "+defaultItem , "Pass");
					            Reusable.takeScreenshot(driver, object_value,TC_Name);
				    				
				    		}
				    		else
				    		{
				    			ExcelUtils.setCellData(val, iTestCaseRow, val_Uname);		    		
					            Log.info(Edit_Username+"ExportExcel Done For "+Edit_Username );
					            Extend_Report.AddReport( "ExportExcel Done For  "+Edit_Username.toString() , "Pass");
					            Reusable.takeScreenshot(driver, object_value,TC_Name);
				    		}
				    		
				    		}
				    	}
					}
			    	
			    	catch(Exception e)
					{	
			    		Reusable.takeScreenshot(driver, object_value,TC_Name);
			    		ExcelUtils.setExcelFile(Constant.Path_Excel+Constant.File_RunManager, "Run_Manager");		 
			    		int val_run=ExcelUtils.getRowContains_manager(TC_Name, 0, "Run_Manager");		 
			    		ExcelUtils.setCellData_Manager("FAIL", val_run, 2);
			    		Constant.TC_Status= "FAIL";    
			    		ExcelUtils.setExcelFile(Constant.Path_Excel+Constant.File_RunManager, "Run_Manager");			   		 
			    		int val_run2=ExcelUtils.getRowContains_manager(TC_Name, 0, "Run_Manager");			   		 
			   		 	String status=ExcelUtils.getCellData(val_run2, 2);
			   		 	System.out.println("Error in ExportExcel " + status);
			    		Extend_Report.AddReport("Error in  ExportExcel", "Fail");
			    		System.out.println("Error in ExportExcel " + e.toString());
			    		
					}
					
				}
				
				//******************************************************************************************************************************//
				public static void MouseOver_WebElementdblClick(WebDriver driver,int iTestCaseRow,String object_value, String Excel_Workbook,String TC_Name) throws Exception
				{
				
						
					
					try
					{
						
				    	         	
				    	int val_Uname=ExcelUtils.get_column(object_value,Excel_Workbook); 				    	
				    	String sUserName1 = ExcelUtils.getCellData(iTestCaseRow, val_Uname);  
				    	String sUserName = sUserName1.trim();
				    	
				    	
				    	if(sUserName != null && !sUserName.isEmpty()) 
				    	{
				    		Thread.sleep(2000);
				    		WebElement Edit_Username=RepositoryParser.Create_Objects(object_value,Excel_Workbook); 				    		
				    		Actions action = new Actions(driver);
				    		 //Find the targeted element
				    		 
				    		                //Here I used JavascriptExecutor interface to scroll down to the targeted element
				    		 ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView();", Edit_Username);
				    		                //used doubleClick(element) method to do double click action
				    		 action.doubleClick(Edit_Username).build().perform();
				    					    		
				            Log.info(Edit_Username+"MouseOver Operation and double click is done over"+Edit_Username );
				            Extend_Report.AddReport( "MouseOver Operation and double click is done over"+Edit_Username.toString() , "Pass");
				            Reusable.takeScreenshot(driver, object_value,TC_Name);
				            Thread.sleep(1000);
				    	}
					}
			    	
			    	catch(Exception e)
					{	
			    		Reusable.takeScreenshot(driver, object_value,TC_Name);
			    		ExcelUtils.setExcelFile(Constant.Path_Excel+Constant.File_RunManager, "Run_Manager");		 
			    		int val_run=ExcelUtils.getRowContains_manager(TC_Name, 0, "Run_Manager");		 
			    		ExcelUtils.setCellData_Manager("FAIL", val_run, 2);
			    		Constant.TC_Status= "FAIL";    
			    		Extend_Report.AddReport("Error is  MouseOver_WebElement", "Fail");
			    		System.out.println("Error in MouseOver_WebElement " + e.toString());
			    		
					}
					
				}
	//******************************************************************************************************************************//
		@SuppressWarnings("unused")
		public static void WaitForElement(WebDriver driver,int iTestCaseRow,String object_value, String Excel_Workbook,String TC_Name) throws Exception
		{
		
				
			
			try
			{
				
				
				
				
		    	//Edit_Username.clear();            	
		    	int val_Uname=ExcelUtils.get_column(object_value,Excel_Workbook);   
		    	
		    	String sUserName1 = ExcelUtils.getCellData(iTestCaseRow, val_Uname);  
		    	String sUserName = sUserName1.trim();
		    	
		    	if(sUserName != null && !sUserName.isEmpty()) 
		    	{
		    		WebElement Edit_Username=RepositoryParser.Create_Objects(object_value,Excel_Workbook); 	
		    		WebDriverWait wait=new WebDriverWait(driver, 10);
		    	
		    		wait.until(ExpectedConditions.visibilityOfElementLocated(Constant.LocatorV));
		    		Extend_Report.AddReport("Sucess in waiting for Element :" + Edit_Username, "Pass");
		            
		    	}
		    	
		    	Reusable.takeScreenshot(driver, object_value,TC_Name);
			}
	    	
	    	catch(Exception e)
			{	
	    		Reusable.takeScreenshot(driver, object_value,TC_Name);
	    		ExcelUtils.setExcelFile(Constant.Path_Excel+Constant.File_RunManager, "Run_Manager");		 
	    		int val_run=ExcelUtils.getRowContains_manager(TC_Name, 0, "Run_Manager");		 
	    		ExcelUtils.setCellData_Manager("FAIL", val_run, 2);
	    		Constant.TC_Status= "FAIL";    
	    		 Extend_Report.AddReport("Error in WaitForElement", "Fail");
	    		System.out.println("Error in WaitForElement " + e.toString());
	    		Log.info(" Error in WaitForElement" );
			}
			
		}
		//******************************************************************************************************************************//
				@SuppressWarnings("unused")
				public static void SikuliClick(WebDriver driver,int iTestCaseRow,String object_value, String Excel_Workbook,String TC_Name) throws Exception
				{
				
						
					
					try
					{
						
						
						
						
				    	//Edit_Username.clear();            	
				    	int val_Uname=ExcelUtils.get_column(object_value,Excel_Workbook);   
				    	
				    	String sUserName1 = ExcelUtils.getCellData(iTestCaseRow, val_Uname);  
				    	String sUserName = sUserName1.trim();
				    	
				    	System.out.println(sUserName.trim());
				    	
				    	if(sUserName != null && !sUserName.isEmpty()) 
				    	{
				    			
				    		String sikulipart = Constant.Sikuli_Path;
				    		
				    		 Screen screen = new Screen();
				    		 
				    		String Carclick = sikulipart.concat(sUserName.trim()+".PNG");
				    		Pattern Carclickp = new Pattern(Carclick);
				    		
				    		 screen.click(Carclickp);    		
				    
				    		
				    		
				    		Extend_Report.AddReport("Success in SikuliClick :" + sUserName, "Pass");
				            
				    	}
				    	
				    	Reusable.takeScreenshot(driver, object_value,TC_Name);
					}
			    	
			    	catch(Exception e)
					{	
			    		Reusable.takeScreenshot(driver, object_value,TC_Name);
			    		ExcelUtils.setExcelFile(Constant.Path_Excel+Constant.File_RunManager, "Run_Manager");		 
			    		int val_run=ExcelUtils.getRowContains_manager(TC_Name, 0, "Run_Manager");		 
			    		ExcelUtils.setCellData_Manager("FAIL", val_run, 2);
			    		 Extend_Report.AddReport("SikuliClick", "Fail");
			    		 Constant.TC_Status= "FAIL";    
			    		System.out.println("SikuliClick " + e.toString());
			    		Log.info(" Error in SikuliClick" );
					}
					
				}
	//******************************************************************************************************************************//
	public static void Click_Button(WebDriver driver,int iTestCaseRow,String object_value, String Excel_Workbook,String TC_Name) throws Exception
	{
		try{
			
			
	    	int val_Uname=ExcelUtils.get_column(object_value,Excel_Workbook);            	
	    	String sUserName1 = ExcelUtils.getCellData(iTestCaseRow, val_Uname);  
	    	String sUserName = sUserName1.trim();
	    	
	    	
			if(sUserName != null && !sUserName.isEmpty()) 
	    	{
	    		
	    		driver.manage().timeouts().implicitlyWait(9, TimeUnit.SECONDS);
	    		WebElement Edit_Username=RepositoryParser.Create_Objects(object_value,Excel_Workbook); 	    		
	    		String idprint = Edit_Username.getAttribute("id");
	    		String nameprint=Edit_Username.getAttribute("name");
	    		Edit_Username.click();	    	
	            Log.info(Edit_Username+" is Clicked " );     
	            String[] partsreport = Edit_Username.toString().split("->");
	            String partsrep=partsreport[1];
	            partsrep=partsrep.replace("[", "");
	            partsrep=partsrep.replace("]", "");
	            partsrep=partsrep.trim();
	        	
	        	Extend_Report.AddReport(partsrep +" Button is Clicked ", "Pass");
	        	 Thread.sleep(1000);
	        
	    	}
	    	Reusable.takeScreenshot(driver, object_value,TC_Name);
			}
		catch(Exception e)
		{
			Reusable.takeScreenshot(driver, object_value,TC_Name);
			ExcelUtils.setExcelFile(Constant.Path_Excel+Constant.File_RunManager, "Run_Manager");		 
    		int val_run=ExcelUtils.getRowContains_manager(TC_Name, 0, "Run_Manager");		 
    		ExcelUtils.setCellData_Manager("FAIL", val_run, 2);
    		
    		int val_run2=ExcelUtils.getRowContains_manager(TC_Name, 0, "Run_Manager");			   		 
   		 	String status=ExcelUtils.getCellData(val_run2, 2);
   		 	System.out.println("Error in ExportExcel " + status);   		 	
    		Extend_Report.AddReport("Error in  ExportExcel", "Fail");    		
    		System.out.println("Error in ExportExcel " + e.toString());
    		Constant.TC_Status= "FAIL";    
    		
			Extend_Report.AddReport("Error in  Click_Button", "Fail");
			Log.info(object_value+" is not Clicked " );
			System.out.println("Error in Click_Button " + e.toString());
		}
    	
    	
		
	}
	
	//******************************************************************************************************************************//
		public static void Click_Button_Optional(WebDriver driver,int iTestCaseRow,String object_value, String Excel_Workbook,String TC_Name) throws Exception
		{
			try{
				
				
		    	int val_Uname=ExcelUtils.get_column(object_value,Excel_Workbook);            	
		    	String sUserName1 = ExcelUtils.getCellData(iTestCaseRow, val_Uname);  
		    	String sUserName = sUserName1.trim();
		    	
		    	
				if(sUserName != null && !sUserName.isEmpty()) 
		    	{
		    		
		    		driver.manage().timeouts().implicitlyWait(12, TimeUnit.SECONDS);
		    		WebElement Edit_Username=RepositoryParser.Create_Objects(object_value,Excel_Workbook); 
		    		
		    		String idprint = Edit_Username.getAttribute("id");
		    		String nameprint=Edit_Username.getAttribute("name");
		    		Edit_Username.click();	    	
		            Log.info(Edit_Username+" is Clicked " );     
		            String[] partsreport = Edit_Username.toString().split("->");
		            String partsrep=partsreport[1];
		            partsrep=partsrep.replace("[", "");
		            partsrep=partsrep.replace("]", "");
		            partsrep=partsrep.trim();
		        	
		        	Extend_Report.AddReport(partsrep +" Button is Clicked ", "Pass");
		        	
		        
		    	}
		    	Reusable.takeScreenshot(driver, object_value,TC_Name);
				}
			catch(Exception e)
			{
				
				Reusable.takeScreenshot(driver, object_value,TC_Name);
				Extend_Report.AddReport("Optional  Button is not clicked", "Pass");
				
			}
	    	
	    	
			
		}
	//******************************************************************************************************************************//
	public static void Select_Value_DropDown(WebDriver driver,int iTestCaseRow,String object_value, String Excel_Workbook,String TC_Name) throws Exception
	{
	
		
		try
		{
			
	    	int val_Uname=ExcelUtils.get_column(object_value,Excel_Workbook);            	
	    	String sUserName1 = ExcelUtils.getCellData(iTestCaseRow, val_Uname);  
	    	String sUserName = sUserName1.trim(); 
	    	if(sUserName != null && !sUserName.isEmpty()) 
	    	{
	    		driver.manage().timeouts().implicitlyWait(12, TimeUnit.SECONDS);
	    		WebElement Edit_Username=RepositoryParser.Create_Objects(object_value,Excel_Workbook); 
	    		
	    		Select select = new Select(Edit_Username);	    		
	    		select.selectByVisibleText(sUserName);
	            Log.info(sUserName+" is entered in text box" );
	            Extend_Report.AddReport( sUserName +" is Selected in the "+ object_value +" DropDown", "Pass");
	            Reusable.takeScreenshot(driver, object_value,TC_Name);
	            Thread.sleep(1000);
	    	}
	    	Log.info(sUserName+" is  Selected " );
		}
    	
    	catch(Exception e)
		{
    		Reusable.takeScreenshot(driver, object_value,TC_Name);
    		ExcelUtils.setExcelFile(Constant.Path_Excel+Constant.File_RunManager, "Run_Manager");		 
    		int val_run=ExcelUtils.getRowContains_manager(TC_Name, 0, "Run_Manager");		 
    		ExcelUtils.setCellData_Manager("FAIL", val_run, 2);
    		
    		int val_run2=ExcelUtils.getRowContains_manager(TC_Name, 0, "Run_Manager");			   		 
   		 	String status=ExcelUtils.getCellData(val_run2, 2);
   		 	System.out.println("Error in ExportExcel " + status);   		 	
    		Extend_Report.AddReport("Error in  ExportExcel", "Fail");    		
    		System.out.println("Error in ExportExcel " + e.toString());
    		Constant.TC_Status= "FAIL";    
    		
    		Extend_Report.AddReport("Error is  Select_Value_DropDown", "Fail");
    		Log.info(" Error is  Select_Value_DropDown " );
    		System.out.println("Error in Select_Value_DropDown " + e.toString());
		}
		
	}
	
	//******************************************************************************************************************************//
		public static void Select_Value_DropDownValue(WebDriver driver,int iTestCaseRow,String object_value, String Excel_Workbook,String TC_Name) throws Exception
		{
		
			
			try
			{
				
		    	int val_Uname=ExcelUtils.get_column(object_value,Excel_Workbook);            	
		    	String sUserName1 = ExcelUtils.getCellData(iTestCaseRow, val_Uname);  
		    	String sUserName = sUserName1.trim(); 
		    	if(sUserName != null && !sUserName.isEmpty()) 
		    	{
		    		
		    		WebElement Edit_Username=RepositoryParser.Create_Objects(object_value,Excel_Workbook); 
		    		Reusable.takeScreenshot((WebDriver) driver, object_value,TC_Name);
		    		Select select = new Select(Edit_Username);		    		
		    		select.selectByValue(sUserName);		
		    		System.out.println("Select_Value_DropDown " + Edit_Username.toString());
		            Log.info(sUserName+" is entered in text box" );
		            Extend_Report.AddReport( sUserName +" is Selected in the "+ object_value +" DropDown", "Pass");
		            Reusable.takeScreenshot(driver, object_value,TC_Name);
		            Thread.sleep(1000);
		    	}
		    	
		    	Log.info(sUserName+" is  Selected " );
			}
	    	
	    	catch(Exception e)
			{
	    		Reusable.takeScreenshot(driver, object_value,TC_Name);
	    		ExcelUtils.setExcelFile(Constant.Path_Excel+Constant.File_RunManager, "Run_Manager");		 
	    		int val_run=ExcelUtils.getRowContains_manager(TC_Name, 0, "Run_Manager");		 
	    		ExcelUtils.setCellData_Manager("FAIL", val_run, 2);
	    		Constant.TC_Status= "FAIL";    
	    		Extend_Report.AddReport("Error is  Select_Value_DropDown", "Fail");
	    		Log.info(" Error is  Select_Value_DropDown " );
	    		System.out.println("Error in Select_Value_DropDown " + e.toString());
			}
			
		}
	
	//******************************************************************************************************************************//
	public static void Check_WebElement(WebDriver driver,int iTestCaseRow,String object_value, String Excel_Workbook,String TC_Name) throws Exception
	{
	
		
		try
		{
			
			
	    	int val_Uname=ExcelUtils.get_column(object_value,Excel_Workbook);            	
	    	String sUserName1 = ExcelUtils.getCellData(iTestCaseRow, val_Uname);  
	    	String sUserName = sUserName1.trim();	    	
	    	if(sUserName != null && !sUserName.isEmpty()) 
	    	{
	    		WebElement Edit_Username=RepositoryParser.Create_Objects(object_value,Excel_Workbook); 
	    		
	    		String val=Edit_Username.getText().trim();
	    		JavascriptExecutor js = (JavascriptExecutor) driver;
	    		js.executeScript("var evt = document.createEvent('MouseEvents');" + "evt.initMouseEvent('click',true, true, window, 0, 0, 0, 0, 0, false, false, false, false, 0,null);" + "arguments[0].dispatchEvent(evt);", Edit_Username);
	    	//	Edit_Username.click();
	    		System.out.println("VAL " + val);
	    		System.out.println("VAL " + sUserName);
	    		if (val.contains(sUserName))
	    		{
	    			System.out.println("PASS CONTAINS");
	    			if (val.length()<80)
	    			{
	    				Log.info(sUserName+" is Present in Page" );
			            Extend_Report.AddReport( val +" is Present in the "+ object_value +" Element ", "Pass");
			            Log.info(val +" is Present in the "+ object_value +" Element " );
			            Reusable.takeScreenshot(driver, object_value,TC_Name);
	    			}
	    			
	    			else
	    			{
	    				Log.info(sUserName+" is Present in Page" );
			            Extend_Report.AddReport( Edit_Username.toString() +" is Present in the "+ object_value +" Element ", "Pass");
			            Log.info(val +" is Present in the "+ object_value +" Element " );
			            Reusable.takeScreenshot(driver, object_value,TC_Name);
	    			}
	    			
	    		}
	    		else
	    		{
	    			System.out.println("FAIL CONTAINS");
	    			ExcelUtils.setExcelFile(Constant.Path_Excel+Constant.File_RunManager, "Run_Manager");		 
	        		int val_run=ExcelUtils.getRowContains_manager(TC_Name, 0, "Run_Manager");		 
	        		ExcelUtils.setCellData_Manager("FAIL", val_run, 2);
	        		Constant.TC_Status= "FAIL";    
	        		Extend_Report.AddReport( val +" is NOT Present in the "+ object_value +" Element ", "Fail");
	        		Reusable.takeScreenshot(driver, object_value,TC_Name);
	    		}
	            
	    		Reusable.takeScreenshot(driver, object_value,TC_Name);
	    	}
		}
    	
    	catch(Exception e)
		{
    		Reusable.takeScreenshot(driver, object_value,TC_Name);
    		ExcelUtils.setExcelFile(Constant.Path_Excel+Constant.File_RunManager, "Run_Manager");		 
    		int val_run=ExcelUtils.getRowContains_manager(TC_Name, 0, "Run_Manager");		 
    		ExcelUtils.setCellData_Manager("FAIL", val_run, 2);
    		Constant.TC_Status= "FAIL";    
    		Extend_Report.AddReport("Error is  Check_WebElement :" +e.toString(), "Fail");
    		
		}
		
	}
	
	//******************************************************************************************************************************//
		public static void ClickLinkText(WebDriver driver,int iTestCaseRow,String object_value, String Excel_Workbook,String TC_Name) throws Exception
		{
		
			
			try
			{
				
				
		    	int val_Uname=ExcelUtils.get_column(object_value,Excel_Workbook);            	
		    	String sUserName1 = ExcelUtils.getCellData(iTestCaseRow, val_Uname);  
		    	String sUserName = sUserName1.trim();	    	
		    	if(sUserName != null && !sUserName.isEmpty()) 
		    	{
		    		//WebElement Edit_Username=RepositoryParser.Create_Objects(object_value,Excel_Workbook); 
		    		WebDriverWait wait = new WebDriverWait(driver,30);
         			//wait.until(ExpectedConditions.visibilityOfElementLocated(By.id(locatorValue));
         			wait.until(ExpectedConditions.visibilityOfElementLocated(By.linkText(sUserName)));             			
         			By locator = By.linkText(sUserName);
         			WebElement linke = (WebElement) driver.findElement(locator);
		    		String val=linke.getText();
		    		JavascriptExecutor js = (JavascriptExecutor) driver;
		    		js.executeScript("var evt = document.createEvent('MouseEvents');" + "evt.initMouseEvent('click',true, true, window, 0, 0, 0, 0, 0, false, false, false, false, 0,null);" + "arguments[0].dispatchEvent(evt);", linke);
		    	//	Edit_Username.click();
		            
		    		Reusable.takeScreenshot(driver, object_value,TC_Name);
		    	}
			}
	    	
	    	catch(Exception e)
			{
	    		Reusable.takeScreenshot(driver, object_value,TC_Name);
	    		ExcelUtils.setExcelFile(Constant.Path_Excel+Constant.File_RunManager, "Run_Manager");		 
	    		int val_run=ExcelUtils.getRowContains_manager(TC_Name, 0, "Run_Manager");		 
	    		ExcelUtils.setCellData_Manager("FAIL", val_run, 2);
	    		Constant.TC_Status= "FAIL";    
	    		Extend_Report.AddReport("Error in  ClickLinkText :" +e.toString(), "Fail");
	    		System.out.println("Error in ClickLinkText"+e.toString());
			}
			
		}
	//******************************************************************************************************************************//
	public static void Click_CheckBox(WebDriver driver,int iTestCaseRow,String object_value, String Excel_Workbook,String TC_Name) throws Exception
	{
	
		
		try
		{
			
			
	    	int val_Uname=ExcelUtils.get_column(object_value,Excel_Workbook);            	
	    	String sUserName1 = ExcelUtils.getCellData(iTestCaseRow, val_Uname);  
	    	String sUserName = sUserName1.trim();	    	
	    	if(sUserName != null && !sUserName.isEmpty()) 
	    	{
	    		WebElement Edit_Username=RepositoryParser.Create_Objects(object_value,Excel_Workbook); 
	    		
	    		
	    		if ( !Edit_Username.isSelected() )
	    		{
	    				Edit_Username.click();
	    				Log.info(Edit_Username +" is Clicked in the Check_Box " );
	    		}
	    		
	    		Reusable.takeScreenshot(driver, object_value,TC_Name);
	    		Extend_Report.AddReport("CheckBox "+ object_value +"is Selected", "Pass");	
	            
	    	}
		}
    	
    	catch(Exception e)
		{
    		Reusable.takeScreenshot(driver, object_value,TC_Name);
    		ExcelUtils.setExcelFile(Constant.Path_Excel+Constant.File_RunManager, "Run_Manager");		 
    		int val_run=ExcelUtils.getRowContains_manager(TC_Name, 0, "Run_Manager");		 
    		ExcelUtils.setCellData_Manager("FAIL", val_run, 2);
    		Constant.TC_Status= "FAIL";    
    		Extend_Report.AddReport("Error is  Click_CheckBox :"+e.toString(), "Fail");
    		
		}
		
	}
	//******************************************************************************************************************************//
	public static void Click_RadioButton(WebDriver driver,int iTestCaseRow,String object_value, String Excel_Workbook,String TC_Name) throws Exception
	{
	
		
		try
		{
			
			
	    	int val_Uname=ExcelUtils.get_column(object_value,Excel_Workbook);            	
	    	String sUserName1 = ExcelUtils.getCellData(iTestCaseRow, val_Uname);  
	    	String sUserName = sUserName1.trim();	    	
	    	if(sUserName != null && !sUserName.isEmpty()) 
	    	{
	    		WebElement Edit_Username=RepositoryParser.Create_Objects(object_value,Excel_Workbook); 
	    		
	    		
	    		if ( !Edit_Username.isSelected() )
	    		{
	    				Edit_Username.click();
	    				Log.info(Edit_Username +" is Clicked in the RadioButton " );
	    		}
	    		Reusable.takeScreenshot(driver, object_value,TC_Name);
	    		Extend_Report.AddReport("RadioButton "+ object_value +"is Selected", "Pass");
	    	}
		}
    	
    	catch(Exception e)
		{
    		Reusable.takeScreenshot(driver, object_value,TC_Name);
    		ExcelUtils.setExcelFile(Constant.Path_Excel+Constant.File_RunManager, "Run_Manager");		 
    		int val_run=ExcelUtils.getRowContains_manager(TC_Name, 0, "Run_Manager");		 
    		ExcelUtils.setCellData_Manager("FAIL", val_run, 2);
    		Constant.TC_Status= "FAIL";    
    		Extend_Report.AddReport("Error is  Click_RadioButton :"+e.toString(), "Fail");
    		
		}
		
	}
	//******************************************************************************************************************************//
	public static void takeScreenshot(WebDriver driver,String sTestCaseName,String TC_Name) throws Exception{
		try{
			int count = 1;		
			
			if (Constant.Global_BrowserType.equalsIgnoreCase("IE"))
				{
					
				try {
					
					String TC = Constant.TC_Name+sTestCaseName+".png";
					
					//System.out.println(TC);
					File tmpDir = new File(TC);
					
					if (tmpDir.exists())
					{
						
						String TC2 = Constant.TC_Name+sTestCaseName+count+".png";
						
						//FileUtils.copyFile(scrFile, new File(TC2));
						count=count+1;
						BufferedImage image = new Robot().createScreenCapture(new Rectangle(Toolkit.getDefaultToolkit().getScreenSize()));

					    ImageIO.write(image, "png", new File(TC2));            

					}
					else
					{
					//	FileUtils.copyFile(scrFile, new File(TC));
						BufferedImage image = new Robot().createScreenCapture(new Rectangle(Toolkit.getDefaultToolkit().getScreenSize()));

					    ImageIO.write(image, "png", new File(TC));     
					}
				    
				} 
				catch (Exception e) {
				    // TODO Auto-generated catch block
				    e.printStackTrace();
				}
				
				}
			else
			{
				File scrFile = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
				// Now you can do whatever you need to do with it, for example copy somewhere
				//String Path=System.getProperty("user.dir") +"//main//java//src//Screenshots//";		
				
				String TC = Constant.TC_Name+sTestCaseName+".png";
				
				//System.out.println(TC);
				File tmpDir = new File(TC);
				
				if (tmpDir.exists())
				{
					
					String TC2 = Constant.TC_Name+sTestCaseName+count+".png";
					
					FileUtils.copyFile(scrFile, new File(TC2));
					count=count+1;
				}
				else
				{
					FileUtils.copyFile(scrFile, new File(TC));
				}
			}
			
		} catch (Exception e){
			
			System.out.println("Screen Shot"+e.toString());
			Extend_Report.AddReport("Error in  takeScreenshot", "Fail");
			Log.error("Class Utils | Method takeScreenshot | Exception occured while capturing ScreenShot : "+e.getMessage());
			//throw new Exception();
		}
	}
	
//********************************** IE flickering**************************************************************************	
	//IE screenprint
	
	
	public static boolean TakeScreenshot(String filePath){
		boolean b = false;
		try {
			
		
		    BufferedImage image = new Robot().createScreenCapture(new Rectangle(Toolkit.getDefaultToolkit().getScreenSize()));

		    b = ImageIO.write(image, "png", new File(filePath));            

		} 
		catch (Exception e) {
		    // TODO Auto-generated catch block
		    e.printStackTrace();
		}
		return b;
	}
//******************************************************************************************************************************//
	
	public static void AysnWait(WebDriver driver,int iTestCaseRow,String object_value, String Excel_Workbook,String TC_Name) throws Exception
	{

			
		
		try
		{
		
			int val_Uname=ExcelUtils.get_column(object_value,Excel_Workbook);            	
	    	String sUserName1 = ExcelUtils.getCellData(iTestCaseRow, val_Uname);  
	    	String sUserName = sUserName1.trim();	    	
	    	if(sUserName != null && !sUserName.isEmpty()) 
	    	{
	    		Thread.sleep(10000);
	    		Reusable.takeScreenshot(driver, object_value,TC_Name);
	    	}
		}
		
		catch(Exception e)
		{	
			Reusable.takeScreenshot(driver, object_value,TC_Name);
			ExcelUtils.setExcelFile(Constant.Path_Excel+Constant.File_RunManager, "Run_Manager");		 
			int val_run=ExcelUtils.getRowContains_manager(TC_Name, 0, "Run_Manager");		 
			ExcelUtils.setCellData_Manager("FAIL", val_run, 2);
			Constant.TC_Status= "FAIL";    
			 Extend_Report.AddReport("Error is  Enter_Value_WebElement", "Fail");
			System.out.println("Error in Enter_Value_WebElement " + e.toString());
			Log.info(" Error in entered in text box" );
		}
		
	}	
	
	//******************************************************************************************************************************//
	
		public static void VerifyEmail(WebDriver driver,int iTestCaseRow,String object_value, String Excel_Workbook,String TC_Name) throws Exception
		{

				
			
			try
			{
			
				int val_Uname=ExcelUtils.get_column(object_value,Excel_Workbook);            	
		    	String sUserName1 = ExcelUtils.getCellData(iTestCaseRow, val_Uname);  
		    	String sUserName = sUserName1.trim();	    	
		    	if(sUserName != null && !sUserName.isEmpty()) 
		    	{
		    		Thread.sleep(50000);
		    		Folder folder = null;
				    Store store = null;
				    System.out.println("***Reading mailbox...");
				    try {
				      Properties props = new Properties();
				      props.put("mail.store.protocol", "imaps");
				      Session session = Session.getInstance(props);
				      store = session.getStore("imaps");
				      store.connect("imap.outlook.com", Constant.Outlook_EmailID, Constant.Outlook_EmailPwd);
				      folder = store.getFolder("Audatex Mails");
				      folder.open(Folder.READ_WRITE);
				      Flags seen = new Flags(Flags.Flag.SEEN);
				      FlagTerm unseenFlagTerm = new FlagTerm(seen, false);
				      Message messages[] = folder.search(unseenFlagTerm);
				     // Message[] messages = folder.getMessages();
				      System.out.println("No of Messages : " + folder.getMessageCount());
				    
				      int count=0;
				      int unread=folder.getUnreadMessageCount();
				      for (int i=0; i < unread; i++)
				      
				      {
				    	  	System.out.println("No of Unread Messages : " + unread);
					        System.out.println("Reading MESSAGE # " + (i + 1) + "...");
					        Message msg = messages[i];
					        String strMailSubject ="", strMailBody ="";
					        //Getting mail subject
					        Object subject = msg.getSubject();
					        System.out.println(subject.toString());
					        //Getting mail body
					        Object content = msg.getContent().toString();
					        Date dtsendval=msg.getSentDate();
					        
					        //Casting objects of mail subject and body into String
					        
					            String body = (String)content;  
					            strMailSubject = (String)subject;
						        strMailBody = (String)content;
						        //Printing mail subject and body
						        System.out.println(strMailSubject);
						        System.out.println(strMailBody);
						        System.out.println(sUserName);
						        System.out.println(dtsendval);
						        if (strMailSubject.contains(sUserName))
						        {
					        	 System.out.println("Found"+strMailSubject);
							     System.out.println("Found"+strMailBody);
							     System.out.println("Found"+sUserName);
							     System.out.println(dtsendval);
							     msg.setFlag(Flag.SEEN, true);
						       	 Extend_Report.AddReport("Outlook verification Subject Success " +strMailSubject , "Pass");
						     	 
						     	 Extend_Report.AddReport("Outlook verification TimeFrame Success " +dtsendval , "Pass");
						     	 count=1;
						     	 return;
						        }
						        else
						        {
						        	msg.setFlag(Flag.SEEN, true);
						        //	unread=folder.getUnreadMessageCount();
						       	
						        }
				       				       
				      }
				      
				      if (count==0)
				       {
				    	  Extend_Report.AddReport("Error in Outlook verification as email not triggered", "Fail");
				    	  ExcelUtils.setExcelFile(Constant.Path_Excel+Constant.File_RunManager, "Run_Manager");		 
							int val_run=ExcelUtils.getRowContains_manager(TC_Name, 0, "Run_Manager");		 
							ExcelUtils.setCellData_Manager("FAIL", val_run, 2);
							Constant.TC_Status= "FAIL";    
				       }
				      
				    }catch(MessagingException messagingException){
				    	messagingException.printStackTrace();
				    	ExcelUtils.setExcelFile(Constant.Path_Excel+Constant.File_RunManager, "Run_Manager");		 
						int val_run=ExcelUtils.getRowContains_manager(TC_Name, 0, "Run_Manager");		 
						ExcelUtils.setCellData_Manager("FAIL", val_run, 2);
						Extend_Report.AddReport("Error in Outlook connection"+ messagingException.toString(), "Fail");
						 Constant.TC_Status= "FAIL";    
						Log.info(" Error in Outlook verification" );
				    }catch(IOException ioException){
				    	ioException.printStackTrace();
				    	ExcelUtils.setExcelFile(Constant.Path_Excel+Constant.File_RunManager, "Run_Manager");		 
						int val_run=ExcelUtils.getRowContains_manager(TC_Name, 0, "Run_Manager");		 
						ExcelUtils.setCellData_Manager("FAIL", val_run, 2);
						Extend_Report.AddReport("Error in Outlook connection"+ ioException.toString(), "Fail");
						 Constant.TC_Status= "FAIL";    
						Log.info(" Error in Outlook verification" );
				    }finally {
				      if (folder != null) {
				    	  try {
							folder.close(true);
						} catch (MessagingException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
							ExcelUtils.setExcelFile(Constant.Path_Excel+Constant.File_RunManager, "Run_Manager");		 
							int val_run=ExcelUtils.getRowContains_manager(TC_Name, 0, "Run_Manager");		 
							ExcelUtils.setCellData_Manager("FAIL", val_run, 2);
							Constant.TC_Status= "FAIL";    
							 Extend_Report.AddReport("Error in Outlook connection"+ e.toString(), "Fail");
							System.out.println("Error in Outlook connection " + e.toString());
							Log.info(" Error in Outlook connection" );
						}
				      }
				      if (store != null) {
				    	  try {
							store.close();
						} catch (MessagingException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
							ExcelUtils.setExcelFile(Constant.Path_Excel+Constant.File_RunManager, "Run_Manager");		 
							int val_run=ExcelUtils.getRowContains_manager(TC_Name, 0, "Run_Manager");		 
							ExcelUtils.setCellData_Manager("FAIL", val_run, 2);
							Constant.TC_Status= "FAIL";    
							 Extend_Report.AddReport("Error in Outlook connection"+ e.toString(), "Fail");
							System.out.println("Error in Outlook verification " + e.toString());
							Log.info(" Error in Outlook verification" );
						}
				      }
				    }	

		    		Reusable.takeScreenshot(driver, object_value,TC_Name);
		    	}
			}
			
			catch(Exception e)
			{	
				
				ExcelUtils.setExcelFile(Constant.Path_Excel+Constant.File_RunManager, "Run_Manager");		 
				int val_run=ExcelUtils.getRowContains_manager(TC_Name, 0, "Run_Manager");		 
				ExcelUtils.setCellData_Manager("FAIL", val_run, 2);
				Constant.TC_Status= "FAIL";    
				 Extend_Report.AddReport("Error in Outlook connection"+ e.toString(), "Fail");
				System.out.println("Error in Outlook verification " + e.toString());
				Log.info(" Error in Outlook verification" );
			}
			
		}	
	//******************************************************************************************************************************//

	public static void FrameChange(WebDriver driver,int iTestCaseRow,String object_value, String Excel_Workbook,String TC_Name) throws Exception
	{
	
		
		try
		{
			
			
	    	int val_Uname=ExcelUtils.get_column(object_value,Excel_Workbook);            	
	    	String sUserName1 = ExcelUtils.getCellData(iTestCaseRow, val_Uname);  
	    	String sUserName = sUserName1.trim();	    	
	    	if(sUserName != null && !sUserName.isEmpty()) 
	    	{
	    		Thread.sleep(4000);
	    		WebElement Edit_Username=RepositoryParser.Create_Objects(object_value,Excel_Workbook); 
	    		
	    		
	    		if ( !Edit_Username.isSelected() )
	    		{
	    			
	    			driver.switchTo().frame(Edit_Username);
	    			System.out.println(Edit_Username.toString());
	    			
	    			Log.info(Edit_Username +" Switched to Frame " );
	    		}
	    		Reusable.takeScreenshot(driver, object_value,TC_Name);
	    		Extend_Report.AddReport("Switched to Frame " +Edit_Username , "Pass");
	    	}
	    	else
	    	{
	    		driver.switchTo().defaultContent();
	    	}
		}
    	
    	catch(Exception e)
		{
    		Reusable.takeScreenshot(driver, object_value,TC_Name);
    		ExcelUtils.setExcelFile(Constant.Path_Excel+Constant.File_RunManager, "Run_Manager");		 
    		int val_run=ExcelUtils.getRowContains_manager(TC_Name, 0, "Run_Manager");		 
    		ExcelUtils.setCellData_Manager("FAIL", val_run, 2);
    		Constant.TC_Status= "FAIL";    
    		Extend_Report.AddReport("Error is  Click_RadioButton :"+e.toString(), "Fail");
    		
		}
		
	}




//******************************************************************************************************************************//
	public static void ErrorValidation(WebDriver driver,int iTestCaseRow,String object_value, String Excel_Workbook,String TC_Name) throws Exception
	{
	
		
		try
		{
			
			
	    	int val_Uname=ExcelUtils.get_column(object_value,Excel_Workbook);            	
	    	String sUserName1 = ExcelUtils.getCellData(iTestCaseRow, val_Uname);  
	    	String sUserName = sUserName1.trim();	    	
	    	if(sUserName != null && !sUserName.isEmpty()) 
	    	{
	    		WebElement Edit_Username=RepositoryParser.Create_ObjectsQuick(object_value,Excel_Workbook); 
	    		
	    		
	    		if ( Edit_Username.isDisplayed() )
	    		{
	    			Extend_Report.AddReport("Error in ErrorValidation ", "Fail");
	    			Extend_Report.AddReport("Error Message  "+ Edit_Username.getText() +" is Displayed", "Fail");	
	    	  		ExcelUtils.setExcelFile(Constant.Path_Excel+Constant.File_RunManager, "Run_Manager");		 
	    	  		int val_run=ExcelUtils.getRowContains_manager(TC_Name, 0, "Run_Manager");		 
	    	  		ExcelUtils.setCellData_Manager("FAIL", val_run, 2);
	    	  		Constant.TC_Status= "FAIL";    
	    	  		
	    		}
	    		else
	    		{
	    			Extend_Report.AddReport("NO Error Message is Displayed", "Pass");	
	    		}
	    		Reusable.takeScreenshot(driver, object_value,TC_Name);
	    		
	            
	    	}
		}
  	
  	catch(Exception e)
		{
  		Reusable.takeScreenshot(driver, object_value,TC_Name);
  		Extend_Report.AddReport("NO Error Message is Displayed", "Pass");
  		
		}
	}
	
	//******************************************************************************************************************************//

		public static void AutoITUpload(WebDriver driver,int iTestCaseRow,String object_value, String Excel_Workbook,String TC_Name) throws Exception
		{
		
			
			try
			{
				
				
		    	int val_Uname=ExcelUtils.get_column(object_value,Excel_Workbook);            	
		    	String sUserName1 = ExcelUtils.getCellData(iTestCaseRow, val_Uname);  
		    	String sUserName = sUserName1.trim();	    	
		    	if(sUserName != null && !sUserName.isEmpty()) 
		    	{
		    		Thread.sleep(2000);
		    	//	WebElement Edit_Username=RepositoryParser.Create_Objects(object_value,Excel_Workbook); 
		    		
		    		
		    		
		    		if( sUserName.equalsIgnoreCase("Quote") || sUserName.equalsIgnoreCase("Receipt"))
		    		{
		    			
		    			String sikulipart = "C:\\Users\\selvakumar.c\\Downloads\\Fram6\\Fram6 - Copy\\src\\main\\java\\Upload\\";			    		
			    		Screen screen = new Screen();			    		 
			    		String File_Edit = sikulipart.concat("FileUpload_Edit.PNG");
			    		String File_UpButton = sikulipart.concat("FileUpload_Button.PNG");
			    		Pattern File_Editp = new Pattern(File_Edit);	
			    		Pattern File_UpButtonp = new Pattern(File_UpButton);
			    		screen.click(File_Edit);  			    		
			    		String STrUpppath=sikulipart+"Quote 1.pdf";
						screen.type(File_Editp,STrUpppath);
						screen.click(File_UpButtonp);

		    			
		    			
		    		/*	int BrowserCOl = ExcelUtils.get_column("Browser_Type","Attachments");	
		    			
		    			String sBrowserName = ExcelUtils.getCellData(iTestCaseRow, BrowserCOl);	
		    			
		    			if   (sBrowserName.equals("Mozilla"))
		    			{
		    				System.out.println(sUserName + sBrowserName);
		    				
		    				Thread.sleep(1000);
		    				try
		    				{
		    					Process p =Runtime.getRuntime().exec(Constant.AutoIT_UploadPath+"\\QuoteUpload_Firefox.exe");
		    					p.waitFor();
		    				}
		    				catch(Exception e)
		    				{
		    					System.out.println("Error in Auto IT IE ");
		    				}
		    			}
		    			else if (sBrowserName.equals("Chrome"))
		    			{
		    				System.out.println(sUserName + sBrowserName);
		    				try
		    				{
		    				Thread.sleep(1000);
		    				Process p =Runtime.getRuntime().exec(Constant.AutoIT_UploadPath+"\\QuoteUpload_Chrome.exe");
		    				p.waitFor();
		    				}
		    				catch(Exception e)
		    				{
		    					System.out.println("Error in Auto IT Mozilla");
		    				}
		    			}
		    			else if (sBrowserName.equals("IE"))
		    			{
		    				System.out.println(sUserName + sBrowserName);
		    				try
		    				{
		    				Thread.sleep(1000);
		    				Process p =Runtime.getRuntime().exec(Constant.AutoIT_UploadPath+"\\QuoteUpload_IE.exe");
		    				p.waitFor();
		    				}
		    				catch(Exception e)
		    				{
		    					System.out.println("Error in Auto IT Chrome");
		    				}
		    			}*/
		    		}
		    		else if (sUserName.equalsIgnoreCase("Invoice"))
		    		{
		    			System.out.println(sUserName);
		    			int BrowserCOl = ExcelUtils.get_column("Login_Browser_Type","Test_Case");	
		    			String sBrowserName = ExcelUtils.getCellData(iTestCaseRow, BrowserCOl);	
		    			if   (sBrowserName.equals("Mozilla"))
		    			{
		    				Thread.sleep(1000);
		    				Runtime.getRuntime().exec(Constant.AutoIT_UploadPath+"\\InvoiceUpload_Firefox.exe");
		    			}
		    			else if (sBrowserName.equals("Chrome"))
		    			{
		    				Thread.sleep(1000);
		    				Runtime.getRuntime().exec(Constant.AutoIT_UploadPath+"\\InvoiceUpload_Chrome.exe");
		    			}
		    			else if (sBrowserName.equals("IE"))
		    			{
		    				Thread.sleep(1000);
		    				Runtime.getRuntime().exec(Constant.AutoIT_UploadPath+"\\InvoiceUpload_IE.exe");
		    			}
		    			
		    		}
		    		else
		    		{
		    			System.out.println("No AUTO IT UPLOAD");
		    		}
		    		Reusable.takeScreenshot(driver, object_value,TC_Name);
		    		Extend_Report.AddReport("Auto IT run for " +sUserName , "Pass");
		    	}
			}
	    	
	    	catch(Exception e)
			{
	    		Reusable.takeScreenshot(driver, object_value,TC_Name);
	    		ExcelUtils.setExcelFile(Constant.Path_Excel+Constant.File_RunManager, "Run_Manager");		 
	    		int val_run=ExcelUtils.getRowContains_manager(TC_Name, 0, "Run_Manager");		 
	    		ExcelUtils.setCellData_Manager("FAIL", val_run, 2);
	    		Constant.TC_Status= "FAIL";    
	    		Extend_Report.AddReport("Auto IT run Failed "+e.toString(), "Fail");
	    		
			}
			
		}
		
		//******************************************************************************************************************************//

				public static WebDriver OpenBrowser(WebDriver driver,int iTestCaseRow,String object_value, String Excel_Workbook,String TC_Name) throws Exception
				{
				
					
					try
					{
							
						String ENV_URL = null;
						
				    	int val_Uname=ExcelUtils.get_column(object_value,Excel_Workbook);            	
				    	String sUserName1 = ExcelUtils.getCellData(iTestCaseRow, val_Uname);  
				    	String sUserName = sUserName1.trim();	    	
				    	if(sUserName != null && !sUserName.isEmpty()) 
				    	{
				    		Thread.sleep(2000);
				    		//WebElement Edit_Username=RepositoryParser.Create_Objects(object_value,Excel_Workbook); 
				    		//System.out.println(Edit_Username);
				    		 int BrowserENV = ExcelUtils.get_column("Login_ENV","Login");
				    		 int BrowserCOl = ExcelUtils.get_column("Login_Browser_Type","Login");	
				    		 
				 		    String environe = ExcelUtils.getCellData(iTestCaseRow, BrowserENV);
				 		 
				 			if (environe.equalsIgnoreCase("Sys"))
				 		    {
				 		    	
				 		    	ENV_URL=Constant.SYS_URL;
				 		    	
				 		    }
				 		    else if(environe.equalsIgnoreCase("Dev"))
				 		    {
				 		    	
				 		    	ENV_URL=Constant.DEV_URL;
				 		    }
				 		    
				 		    else if(environe.equalsIgnoreCase("UAT"))
				 		    {
				 		    	ENV_URL=Constant.UAT_URL;
				 		    }
				 		try{
				 		String sBrowserName = ExcelUtils.getCellData(iTestCaseRow, BrowserCOl);	
				 		Constant.Global_BrowserType=sBrowserName.trim();
				 		if(sBrowserName.equals("IE")){
				 			//Utils.takeScreenshot(driver, "Browser_Invoke");
				 			//System.setProperty("webdriver.gecko.driver", "C:\\Users\\selvakumar.c\\Desktop\\Selenium\\jar\\geckodriver-v0.18.0-win64\\geckodriver.exe");
				 			//driver = new FirefoxDriver();
				 			System.out.println("Internet Explorer is selected");
				 			System.setProperty("webdriver.ie.driver",Constant.Path_IEDriver);
				 			DesiredCapabilities capabilities = DesiredCapabilities.internetExplorer();
							capabilities.setCapability(InternetExplorerDriver.INITIAL_BROWSER_URL,Constant.SYS_URL);							//
							capabilities.setCapability(InternetExplorerDriver.INTRODUCE_FLAKINESS_BY_IGNORING_SECURITY_DOMAINS, true);
							capabilities.setCapability(InternetExplorerDriver.IE_ENSURE_CLEAN_SESSION, true);
							capabilities.setCapability("requireWindowFocus", true);
							capabilities.setCapability("ignoreZoomSetting", true); 
							//capabilities.setCapability("nativeEvents", false);    
							capabilities.setCapability("unexpectedAlertBehaviour", "accept");
							capabilities.setCapability("ignoreProtectedModeSettings", true);
							capabilities.setCapability("disable-popup-blocking", true);
							capabilities.setCapability("enablePersistentHover", true);
							driver = new InternetExplorerDriver(capabilities);
				 			Log.info("New driver instantiated");
				 		    driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
				 		    Log.info("Implicit wait applied on the driver for 10 seconds");
				 		    //driver.get("https://www.google.com.au/webhp?ie=utf-8&oe=utf-8&gws_rd=cr&ei=VfZvWYvnAoG80ATDsaTQAQ");
				 		 //  driver.manage().deleteAllCookies();
				 		   
				 		  // driver.get(ENV_URL);
				 		    Extend_Report.AddReport("URL Launched "+ENV_URL + "in IE Browser" , "Pass");
				 		    Log.info("Web application launched successfully");
				 			}
				 		
				 		else if (sBrowserName.equals("Mozilla")){
				 			//Utils.takeScreenshot(driver, "Browser_Invoke");
				 			System.setProperty("webdriver.gecko.driver", Constant.Path_FireFoxDriver);
				 			FirefoxProfile profile = new FirefoxProfile();
				 			profile.setPreference("intl.accept_languages","en-AU");
				 			System.out.println("Mozilla  Explorer is selected");			
				 			driver = new FirefoxDriver();
				 			Log.info("New driver instantiated");
				 		    driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
				 		    Log.info("Implicit wait applied on the driver for 10 seconds");
				 		    driver. manage(). window().maximize();
				 		    driver.get(ENV_URL);
				 		    Extend_Report.AddReport("URL Launched "+ENV_URL + "in Mozilla Firefox Browser" , "Pass");
				 		    Log.info("Web application launched successfully");
				 			
				 		}
				 		else if (sBrowserName.equals("Chrome")){
				 			//Utils.takeScreenshot(driver, "Browser_Invoke");
				 			System.setProperty("webdriver.chrome.driver", Constant.Path_ChromeDriver);				 			
				 			ChromeOptions options = new ChromeOptions();				 			
				 			options.addArguments("--lang=en-AU");	
				 			Map<String, Object> prefs = new HashMap<String, Object>();
				 			prefs.put("intl.accept_languages", "en-AU");
				 			options.setExperimentalOption("prefs", prefs);
				 			driver = new ChromeDriver(options);				 			
				 			System.out.println("Chrome Browser is selected");				 			
				 			Log.info("New driver instantiated");
				 		    driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
				 		    Log.info("Implicit wait applied on the driver for 10 seconds");
				 		    driver. manage(). window().maximize();
				 		    driver.get(ENV_URL);	
				 		 
				 		  
				 		    Extend_Report.AddReport("URL Launched "+ENV_URL + "in Chrome Browser" , "Pass");
				 		    Log.info("Web application launched successfully");
				 			
				 		}
				 		if(sBrowserName.equals("EDGE")){
				 			//Utils.takeScreenshot(driver, "Browser_Invoke");
				 			//System.setProperty("webdriver.gecko.driver", "C:\\Users\\selvakumar.c\\Desktop\\Selenium\\jar\\geckodriver-v0.18.0-win64\\geckodriver.exe");
				 			//driver = new FirefoxDriver();
				 			System.out.println("EDGE Explorer is selected");
				 			System.setProperty("webdriver.edge.driver",Constant.Path_EDGEDriver);				 			
							driver = new EdgeDriver();
				 			Log.info("New driver instantiated");
				 		    driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
				 		    Log.info("Implicit wait applied on the driver for 10 seconds");
				 		    //driver.get("https://www.google.com.au/webhp?ie=utf-8&oe=utf-8&gws_rd=cr&ei=VfZvWYvnAoG80ATDsaTQAQ");
				 		 //  driver.manage().deleteAllCookies();
				 		   
				 		  // driver.get(ENV_URL);
				 		    Extend_Report.AddReport("URL Launched "+ENV_URL + "in IE Browser" , "Pass");
				 		    Log.info("Web application launched successfully");
				 			}
				 		}catch (Exception e){
				 			//Utils.takeScreenshot(driver, "Browser_Invoke");
				 			System.out.println("Browser Error");
				 			Log.error("Class Utils | Method OpenBrowser | Exception desc : "+e.getMessage());
				 		}
				 		
				 		new BaseClass(driver);	
				    		
				    		
				    	}
					}
			    	
			    	catch(Exception e)
					{
			    		Reusable.takeScreenshot(driver, object_value,TC_Name);
			    		ExcelUtils.setExcelFile(Constant.Path_Excel+Constant.File_RunManager, "Run_Manager");		 
			    		int val_run=ExcelUtils.getRowContains_manager(TC_Name, 0, "Run_Manager");		 
			    		ExcelUtils.setCellData_Manager("FAIL", val_run, 2);
			    		Constant.TC_Status= "FAIL";    
			    		Extend_Report.AddReport("OpenBrowser Failed "+e.toString(), "Fail");
			    		
					}
					return driver;
					
				}
				
				//******************************************************************************************************************************//

				public static void closeBrowser(WebDriver driver,int iTestCaseRow,String object_value, String Excel_Workbook,String TC_Name) throws Exception
				{
				
					
					try
					{
						
						
				    	int val_Uname=ExcelUtils.get_column(object_value,Excel_Workbook);            	
				    	String sUserName1 = ExcelUtils.getCellData(iTestCaseRow, val_Uname);  
				    	String sUserName = sUserName1.trim();	    	
				    	if(sUserName != null && !sUserName.isEmpty()) 
				    	{
				    		Thread.sleep(2000);
				    	//	WebElement Edit_Username=RepositoryParser.Create_Objects(object_value,Excel_Workbook); 
				    		
				    		if (driver!=null)
							  {
					  			driver.quit();
							  }	 
					    
				    		
				    	}
					}
			    	
			    	catch(Exception e)
					{
			    		Reusable.takeScreenshot(driver, object_value,TC_Name);
			    		ExcelUtils.setExcelFile(Constant.Path_Excel+Constant.File_RunManager, "Run_Manager");		 
			    		int val_run=ExcelUtils.getRowContains_manager(TC_Name, 0, "Run_Manager");		 
			    		ExcelUtils.setCellData_Manager("FAIL", val_run, 2);
			    		Constant.TC_Status= "FAIL";    
			    		Extend_Report.AddReport("closeBrowser Failed "+e.toString(), "Fail");
			    		
					}
					
				}
				
				//******************************************************************************************************************************//
				public static void AysnWaitBrowser(WebDriver driver,int iTestCaseRow,String object_value, String Excel_Workbook,String TC_Name) throws Exception
				{
				
						
					
					try
					{
						
				    	int val_Uname=ExcelUtils.get_column(object_value,Excel_Workbook);
				    	
				    	String sUserName1 = ExcelUtils.getCellData(iTestCaseRow, val_Uname);  
				    	
				    	String sUserName = sUserName1.trim();
				    	
				    	
				    	
				    	if(sUserName != null && !sUserName.isEmpty()) 
				    	{
				    		Thread.sleep(100000);				    		
				    		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
				            Log.info("Browser explicitly waited for 400 seconds" );	           
				            Extend_Report.AddReport( "Browser explicitly waited for 400 seconds", "Pass");
				            Reusable.takeScreenshot(driver, object_value,TC_Name);
				            
				    	}
					}
			    	
			    	catch(Exception e)
					{	
			    		Reusable.takeScreenshot(driver, object_value,TC_Name);
			    		ExcelUtils.setExcelFile(Constant.Path_Excel+Constant.File_RunManager, "Run_Manager");		 
			    		int val_run=ExcelUtils.getRowContains_manager(TC_Name, 0, "Run_Manager");		 
			    		ExcelUtils.setCellData_Manager("FAIL", val_run, 2);
			    		Constant.TC_Status= "FAIL";    
			    		 Extend_Report.AddReport("Error in AysnWaitBrowser", "Fail");
			    		System.out.println("Error in AysnWaitBrowser " + e.toString());
			    		Log.info(" Error in AysnWaitBrowser" );
					}
					
				}
				
				//******************************************************************************************************************************//
				public static void MouseOver_WebElementClickRandom(WebDriver driver,int iTestCaseRow,String object_value, String Excel_Workbook,String TC_Name) throws Exception
				{
				
						
					
					try
					{
						
				    	         	
				    	int val_Uname=ExcelUtils.get_column(object_value,Excel_Workbook); 				    	
				    	String sUserName1 = ExcelUtils.getCellData(iTestCaseRow, val_Uname);  
				    	String sUserName = sUserName1.trim();
				    	
				    	
				    	if(sUserName != null && !sUserName.isEmpty()) 
				    	{
				    	/*	WebElement Edit_Username=RepositoryParser.Create_Objects(object_value,Excel_Workbook); 				    		
				    		Actions action = new Actions(driver);
				    		Point classname = Edit_Username.getLocation();
				            int xcordi = classname.getX();
				            int ycordi = classname.getY();
				            System.out.println(xcordi);
				            System.out.println(ycordi);
				            int ImageWidth = Edit_Username.getSize().getWidth();
				            int ImageHeight = Edit_Username.getSize().getHeight(); 
				            System.out.println(ImageWidth);
				            System.out.println(ImageHeight);
				        //    JavascriptExecutor js = (JavascriptExecutor) driver;
				    		//js.executeScript("var evt = document.createEvent('MouseEvents');" + "evt.initMouseEvent('click',true, true, window, 400, 300, 200, 100, 0, false, false, false, false, 0,null);" + "arguments[0].dispatchEvent(evt);", Edit_Username);
				            JavascriptExecutor executor = (JavascriptExecutor) driver;
				            executor.executeScript("window.scroll(" + xcordi + ", " + ycordi + ");");
				            executor.executeScript("arguments[0].click();", Edit_Username);
				            Log.info(Edit_Username+"MouseOver Operation and click is done over"+Edit_Username );
				            Extend_Report.AddReport( "MouseOver Operation and click is done over"+Edit_Username.toString() , "Pass");
				            Reusable.takeScreenshot(driver, object_value,TC_Name);*/
				    		
				    		
				    		
				    		
				    		
				            
				    	}
					}
			    	
			    	catch(Exception e)
					{	
			    		Reusable.takeScreenshot(driver, object_value,TC_Name);
			    		ExcelUtils.setExcelFile(Constant.Path_Excel+Constant.File_RunManager, "Run_Manager");		 
			    		int val_run=ExcelUtils.getRowContains_manager(TC_Name, 0, "Run_Manager");		 
			    		ExcelUtils.setCellData_Manager("FAIL", val_run, 2);
			    		Constant.TC_Status= "FAIL";    
			    		Extend_Report.AddReport("Error is  MouseOver_WebElement", "Fail");
			    		System.out.println("Error in MouseOver_WebElement " + e.toString());
			    		
					}
					
				}
	}