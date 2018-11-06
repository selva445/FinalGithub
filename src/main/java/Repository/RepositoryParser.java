package Repository;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Properties;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import pageObjects.BaseClass;
import utility.Constant;
import utility.ExcelUtils;
import utility.Extend_Report;
import utility.Log;

public class RepositoryParser extends BaseClass {

	private static WebElement element = null;
    private static XSSFSheet ExcelWSheet;
    private static XSSFWorkbook ExcelWBook;
    private static XSSFCell Cell;
    private static XSSFRow Row;
	public static Object ele;
 
 //Return Object as  Element using FOR Loop
	public RepositoryParser(WebDriver driver){
    	super(driver);
}     

 public static WebElement Create_Objects(String Logical_Name, String Excel_File) throws Exception{
 	try{
     	
 		String currentDir = System.getProperty("user.dir");        
 		FileInputStream ExcelFile = new FileInputStream(currentDir+"\\src\\main\\java\\Repository\\"+Excel_File+".xlsx");     
         ExcelWBook = new XSSFWorkbook(ExcelFile);
         ExcelWSheet = ExcelWBook.getSheet("Sheet1");
         int RowCount = ExcelWSheet.getLastRowNum();
             		
			
			for (int i=1;i<RowCount+1;i++)
 		{
				try{
					Cell = ExcelWSheet.getRow(i).getCell(0);					
					String CellData = Cell.getStringCellValue();	                    
                 Cell = ExcelWSheet.getRow(i).getCell(1);
                 String CellData1 = Cell.getStringCellValue();     			
             if (CellData1.contains("#"))
             {
            	 if(Logical_Name.equalsIgnoreCase(CellData))
                 {
                 	String locatorType = CellData1.split("#")[0];
             		String locatorValue = CellData1.split("#")[1];
             		
             		 By locator = null;
             		switch(locatorType)
             		{
             		case "Id":
             			
             			WebDriverWait wait = new WebDriverWait(driver,30);
             			//wait.until(ExpectedConditions.visibilityOfElementLocated(By.id(locatorValue));
             			wait.until(ExpectedConditions.visibilityOfElementLocated(By.id(locatorValue)));             			
             			locator = By.id(locatorValue);
             			break;
             		case "Name":
             			WebDriverWait wait1 = new WebDriverWait(driver,30);
             			wait1.until(ExpectedConditions.visibilityOfElementLocated(By.name(locatorValue))); 
             			locator = By.name(locatorValue);
             			
             			break;
             		case "CssSelector":
             			WebDriverWait wait2 = new WebDriverWait(driver,30);
             			wait2.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(locatorValue))); 
             			locator = By.cssSelector(locatorValue);
             			break;
             		case "LinkText":
             			WebDriverWait wait3 = new WebDriverWait(driver,30);
             			wait3.until(ExpectedConditions.visibilityOfElementLocated(By.linkText(locatorValue))); 
             			locator = By.linkText(locatorValue);
             			break;
             		case "PartialLinkText":
             			WebDriverWait wait4 = new WebDriverWait(driver,30);
             			wait4.until(ExpectedConditions.visibilityOfElementLocated(By.partialLinkText(locatorValue))); 
             			locator = By.partialLinkText(locatorValue);
             			break;
             		case "TagName":
             			WebDriverWait wait5 = new WebDriverWait(driver,30);
             			wait5.until(ExpectedConditions.visibilityOfElementLocated(By.tagName(locatorValue))); 
             			locator = By.tagName(locatorValue);
             			break;
             		case "Xpath":
             			WebDriverWait wait6 = new WebDriverWait(driver,30);
             			wait6.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(locatorValue))); 
             			locator = By.xpath(locatorValue);
             			break;
             		}
             		                   	
             			Constant.LocatorV=locator;
	                    WebElement linke = (WebElement) driver.findElement(locator);  
	                	return linke;
                 }
                 
                	 
 		}
       else if (CellData1.contains(":"))
             {
            	 if(Logical_Name.equalsIgnoreCase(CellData))
                 {
                 	String locatorType = CellData1.split(":")[0];
             		String locatorValue = CellData1.split(":")[1];
             		
             		 By locator = null;
             		switch(locatorType)
             		{
             		case "Id":
             			WebDriverWait wait = new WebDriverWait(driver,30);
             			wait.until(ExpectedConditions.visibilityOfElementLocated(By.id(locatorValue)));             			
             			locator = By.id(locatorValue);
             			break;
             		case "Name":
             			WebDriverWait wait1 = new WebDriverWait(driver,30);
             			wait1.until(ExpectedConditions.visibilityOfElementLocated(By.name(locatorValue))); 
             			locator = By.name(locatorValue);
             			break;
             		case "CssSelector":
             			WebDriverWait wait2 = new WebDriverWait(driver,30);
             			wait2.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(locatorValue))); 
             			locator = By.cssSelector(locatorValue);
             			break;
             		case "LinkText":
             			WebDriverWait wait3 = new WebDriverWait(driver,30);
             			wait3.until(ExpectedConditions.visibilityOfElementLocated(By.linkText(locatorValue))); 
             			locator = By.linkText(locatorValue);
             			break;
             		case "PartialLinkText":
             			WebDriverWait wait4 = new WebDriverWait(driver,30);
             			wait4.until(ExpectedConditions.visibilityOfElementLocated(By.partialLinkText(locatorValue))); 
             			locator = By.partialLinkText(locatorValue);
             			break;
             		case "TagName":
             			WebDriverWait wait5 = new WebDriverWait(driver,30);
             			wait5.until(ExpectedConditions.visibilityOfElementLocated(By.tagName(locatorValue))); 
             			locator = By.tagName(locatorValue);
             			break;
             		case "Xpath":
             			WebDriverWait wait6 = new WebDriverWait(driver,30);
             			wait6.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(locatorValue))); 
             			locator = By.xpath(locatorValue);
             			break;
             		}
             		                   	
             			Constant.LocatorV=locator;
	                    WebElement linke = (WebElement) driver.findElement(locator);  
	                	return linke;
                 }
                 
				}
             }
                
				
             catch(Exception e)
				{
            	   
             	System.out.println("Exception in Object Creation " + e.toString());
             	ExcelUtils.setExcelFile(Constant.Path_Excel+Constant.File_RunManager, "Run_Manager");		 
        		int val_run=ExcelUtils.getRowContains_manager(Constant.TC_Name, 0, "Run_Manager");		 
        		ExcelUtils.setCellData_Manager("FAIL", val_run, 2);
        		Extend_Report.AddReport("Error in Object Creation"+e.toString(), "Fail");
        		
				}
				
				
 		}
			
			Log.info("Objects Set for  Page");
			
         
 	}catch (Exception e){
 		
 		Log.error("Objects Not Set for  Page");
 		System.out.println("Exception in Object Creation " + e.toString());
     	ExcelUtils.setExcelFile(Constant.Path_Excel+Constant.File_RunManager, "Run_Manager");		 
		int val_run=ExcelUtils.getRowContains_manager(Constant.TC_Name, 0, "Run_Manager");		 
		ExcelUtils.setCellData_Manager("FAIL", val_run, 2);
		Extend_Report.AddReport("Error in Object Creation"+e.toString(), "Fail");
    		throw(e);
    		}
		return null;
	
    	
 }
 public static Select Create_Objects_Select(String Logical_Name, String Excel_File) throws Exception{
	 	try{
	     	
	 		String currentDir = System.getProperty("user.dir");        
	 		FileInputStream ExcelFile = new FileInputStream(currentDir+"\\src\\main\\java\\Repository\\"+Excel_File+".xlsx");     
	         ExcelWBook = new XSSFWorkbook(ExcelFile);
	         ExcelWSheet = ExcelWBook.getSheet("Sheet1");
	         int RowCount = ExcelWSheet.getLastRowNum();
	             		
				
				for (int i=1;i<RowCount+1;i++)
	 		{
					try{
						Cell = ExcelWSheet.getRow(i).getCell(0);					
						String CellData = Cell.getStringCellValue();	                    
	                 Cell = ExcelWSheet.getRow(i).getCell(1);
	                 String CellData1 = Cell.getStringCellValue();     			
	             if (CellData1.contains("#"))
	             {
	            	 if(Logical_Name.equalsIgnoreCase(CellData))
	                 {
	                 	String locatorType = CellData1.split("#")[0];
	             		String locatorValue = CellData1.split("#")[1];
	             		
	             		 By locator = null;
	             		switch(locatorType)
	             		{
	             		case "Id":
	             			
	             			WebDriverWait wait = new WebDriverWait(driver,30);
	             			//wait.until(ExpectedConditions.visibilityOfElementLocated(By.id(locatorValue));
	             			wait.until(ExpectedConditions.visibilityOfElementLocated(By.id(locatorValue)));             			
	             			locator = By.id(locatorValue);
	             			break;
	             		case "Name":
	             			WebDriverWait wait1 = new WebDriverWait(driver,30);
	             			wait1.until(ExpectedConditions.visibilityOfElementLocated(By.name(locatorValue))); 
	             			locator = By.name(locatorValue);
	             			
	             			break;
	             		case "CssSelector":
	             			WebDriverWait wait2 = new WebDriverWait(driver,30);
	             			wait2.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(locatorValue))); 
	             			locator = By.cssSelector(locatorValue);
	             			break;
	             		case "LinkText":
	             			WebDriverWait wait3 = new WebDriverWait(driver,30);
	             			wait3.until(ExpectedConditions.visibilityOfElementLocated(By.linkText(locatorValue))); 
	             			locator = By.linkText(locatorValue);
	             			break;
	             		case "PartialLinkText":
	             			WebDriverWait wait4 = new WebDriverWait(driver,30);
	             			wait4.until(ExpectedConditions.visibilityOfElementLocated(By.partialLinkText(locatorValue))); 
	             			locator = By.partialLinkText(locatorValue);
	             			break;
	             		case "TagName":
	             			WebDriverWait wait5 = new WebDriverWait(driver,30);
	             			wait5.until(ExpectedConditions.visibilityOfElementLocated(By.tagName(locatorValue))); 
	             			locator = By.tagName(locatorValue);
	             			break;
	             		case "Xpath":
	             			WebDriverWait wait6 = new WebDriverWait(driver,30);
	             			wait6.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(locatorValue))); 
	             			locator = By.xpath(locatorValue);
	             			break;
	             		}
	             		                   	
	             			Constant.LocatorV=locator;
	             			Select linke=new Select(driver.findElement(locator));
		                	return linke;
	                 }
	                 
	                	 
	 		}
	       else if (CellData1.contains(":"))
	             {
	            	 if(Logical_Name.equalsIgnoreCase(CellData))
	                 {
	                 	String locatorType = CellData1.split(":")[0];
	             		String locatorValue = CellData1.split(":")[1];
	             		
	             		 By locator = null;
	             		switch(locatorType)
	             		{
	             		case "Id":
	             			WebDriverWait wait = new WebDriverWait(driver,30);
	             			wait.until(ExpectedConditions.visibilityOfElementLocated(By.id(locatorValue)));             			
	             			locator = By.id(locatorValue);
	             			break;
	             		case "Name":
	             			WebDriverWait wait1 = new WebDriverWait(driver,30);
	             			wait1.until(ExpectedConditions.visibilityOfElementLocated(By.name(locatorValue))); 
	             			locator = By.name(locatorValue);
	             			break;
	             		case "CssSelector":
	             			WebDriverWait wait2 = new WebDriverWait(driver,30);
	             			wait2.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(locatorValue))); 
	             			locator = By.cssSelector(locatorValue);
	             			break;
	             		case "LinkText":
	             			WebDriverWait wait3 = new WebDriverWait(driver,30);
	             			wait3.until(ExpectedConditions.visibilityOfElementLocated(By.linkText(locatorValue))); 
	             			locator = By.linkText(locatorValue);
	             			break;
	             		case "PartialLinkText":
	             			WebDriverWait wait4 = new WebDriverWait(driver,30);
	             			wait4.until(ExpectedConditions.visibilityOfElementLocated(By.partialLinkText(locatorValue))); 
	             			locator = By.partialLinkText(locatorValue);
	             			break;
	             		case "TagName":
	             			WebDriverWait wait5 = new WebDriverWait(driver,30);
	             			wait5.until(ExpectedConditions.visibilityOfElementLocated(By.tagName(locatorValue))); 
	             			locator = By.tagName(locatorValue);
	             			break;
	             		case "Xpath":
	             			WebDriverWait wait6 = new WebDriverWait(driver,30);
	             			wait6.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(locatorValue))); 
	             			locator = By.xpath(locatorValue);
	             			break;
	             		}
	             		                   	
	             			Constant.LocatorV=locator;
	             			 Select linke=new Select(driver.findElement(locator));
	             			// Select linke = driver.findElement(locator);  
			                return linke;
	                 }
	                 
					}
	             }
	                
					
	             catch(Exception e)
					{
	            	   
	             	System.out.println("Exception in Object Creation " + e.toString());
	             	ExcelUtils.setExcelFile(Constant.Path_Excel+Constant.File_RunManager, "Run_Manager");		 
	        		int val_run=ExcelUtils.getRowContains_manager(Constant.TC_Name, 0, "Run_Manager");		 
	        		ExcelUtils.setCellData_Manager("FAIL", val_run, 2);
	        		Extend_Report.AddReport("Error in Object Creation"+e.toString(), "Fail");
	        		
					}
					
					
	 		}
				
				Log.info("Objects Set for  Page");
				
	         
	 	}catch (Exception e){
	 		
	 		Log.error("Objects Not Set for  Page");
	 		System.out.println("Exception in Object Creation " + e.toString());
	     	ExcelUtils.setExcelFile(Constant.Path_Excel+Constant.File_RunManager, "Run_Manager");		 
			int val_run=ExcelUtils.getRowContains_manager(Constant.TC_Name, 0, "Run_Manager");		 
			ExcelUtils.setCellData_Manager("FAIL", val_run, 2);
			Extend_Report.AddReport("Error in Object Creation"+e.toString(), "Fail");
	    		throw(e);
	    		}
			return null;
		
	    	
	 }
 
 public static WebElement Create_ObjectsQuick(String Logical_Name, String Excel_File) throws Exception{
	 	try{
	     	
	 		String currentDir = System.getProperty("user.dir");        
	 		FileInputStream ExcelFile = new FileInputStream(currentDir+"\\src\\main\\java\\Repository\\"+Excel_File+".xlsx");     
	         ExcelWBook = new XSSFWorkbook(ExcelFile);
	         ExcelWSheet = ExcelWBook.getSheet("Sheet1");
	         int RowCount = ExcelWSheet.getLastRowNum();
	             		
				
				for (int i=1;i<RowCount+1;i++)
	 		{
					try{
						Cell = ExcelWSheet.getRow(i).getCell(0);					
						String CellData = Cell.getStringCellValue();	                    
	                 Cell = ExcelWSheet.getRow(i).getCell(1);
	                 String CellData1 = Cell.getStringCellValue();     			
	             if (CellData1.contains("#"))
	             {
	            	 if(Logical_Name.equalsIgnoreCase(CellData))
	                 {
	                 	String locatorType = CellData1.split("#")[0];
	             		String locatorValue = CellData1.split("#")[1];
	             		
	             		 By locator = null;
	             		switch(locatorType)
	             		{
	             		case "Id":
	             			
	             			WebDriverWait wait = new WebDriverWait(driver,2);
	             			//wait.until(ExpectedConditions.visibilityOfElementLocated(By.id(locatorValue));
	             			wait.until(ExpectedConditions.visibilityOfElementLocated(By.id(locatorValue)));             			
	             			locator = By.id(locatorValue);
	             			break;
	             		case "Name":
	             			WebDriverWait wait1 = new WebDriverWait(driver,2);
	             			wait1.until(ExpectedConditions.visibilityOfElementLocated(By.name(locatorValue))); 
	             			locator = By.name(locatorValue);
	             			break;
	             		case "CssSelector":
	             			WebDriverWait wait2 = new WebDriverWait(driver,2);
	             			wait2.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(locatorValue))); 
	             			locator = By.cssSelector(locatorValue);
	             			break;
	             		case "LinkText":
	             			WebDriverWait wait3 = new WebDriverWait(driver,2);
	             			wait3.until(ExpectedConditions.visibilityOfElementLocated(By.linkText(locatorValue))); 
	             			locator = By.linkText(locatorValue);
	             			break;
	             		case "PartialLinkText":
	             			WebDriverWait wait4 = new WebDriverWait(driver,2);
	             			wait4.until(ExpectedConditions.visibilityOfElementLocated(By.partialLinkText(locatorValue))); 
	             			locator = By.partialLinkText(locatorValue);
	             			break;
	             		case "TagName":
	             			WebDriverWait wait5 = new WebDriverWait(driver,2);
	             			wait5.until(ExpectedConditions.visibilityOfElementLocated(By.tagName(locatorValue))); 
	             			locator = By.tagName(locatorValue);
	             			break;
	             		case "Xpath":
	             			WebDriverWait wait6 = new WebDriverWait(driver,2);
	             			wait6.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(locatorValue))); 
	             			locator = By.xpath(locatorValue);
	             			break;
	             		}
	             		                   	
	             			Constant.LocatorV=locator;
		                    WebElement linke = (WebElement) driver.findElement(locator);  
		                	return linke;
	                 }
	                 
	                	 
	 		}
	       else if (CellData1.contains(":"))
	             {
	            	 if(Logical_Name.equalsIgnoreCase(CellData))
	                 {
	                 	String locatorType = CellData1.split(":")[0];
	             		String locatorValue = CellData1.split(":")[1];
	             		
	             		 By locator = null;
	             		switch(locatorType)
	             		{
	             		case "Id":
	             			WebDriverWait wait = new WebDriverWait(driver,2);
	             			wait.until(ExpectedConditions.visibilityOfElementLocated(By.id(locatorValue)));             			
	             			locator = By.id(locatorValue);
	             			break;
	             		case "Name":
	             			WebDriverWait wait1 = new WebDriverWait(driver,2);
	             			wait1.until(ExpectedConditions.visibilityOfElementLocated(By.name(locatorValue))); 
	             			locator = By.name(locatorValue);
	             			break;
	             		case "CssSelector":
	             			WebDriverWait wait2 = new WebDriverWait(driver,2);
	             			wait2.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(locatorValue))); 
	             			locator = By.cssSelector(locatorValue);
	             			break;
	             		case "LinkText":
	             			WebDriverWait wait3 = new WebDriverWait(driver,2);
	             			wait3.until(ExpectedConditions.visibilityOfElementLocated(By.linkText(locatorValue))); 
	             			locator = By.linkText(locatorValue);
	             			break;
	             		case "PartialLinkText":
	             			WebDriverWait wait4 = new WebDriverWait(driver,2);
	             			wait4.until(ExpectedConditions.visibilityOfElementLocated(By.partialLinkText(locatorValue))); 
	             			locator = By.partialLinkText(locatorValue);
	             			break;
	             		case "TagName":
	             			WebDriverWait wait5 = new WebDriverWait(driver,2);
	             			wait5.until(ExpectedConditions.visibilityOfElementLocated(By.tagName(locatorValue))); 
	             			locator = By.tagName(locatorValue);
	             			break;
	             		case "Xpath":
	             			WebDriverWait wait6 = new WebDriverWait(driver,2);
	             			wait6.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(locatorValue))); 
	             			locator = By.xpath(locatorValue);
	             			break;
	             		}
	             		                   	
	             			Constant.LocatorV=locator;
		                    WebElement linke = (WebElement) driver.findElement(locator);  
		                	return linke;
	                 }
	                 
					}
	             }
	                
					
	             catch(Exception e)
					{
	            	   
	             	System.out.println("Exception in Object Creation " + e.toString());
					}
					
					
	 		}
				
				Log.info("Objects Set for  Page");
				
	         
	 	}catch (Exception e){
	 		
	 		Log.error("Objects Not Set for  Page");
	    		throw(e);
	    		}
			return null;
		
	    	
	 }
}