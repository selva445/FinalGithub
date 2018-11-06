package appModules;
import java.io.FileInputStream;
import java.sql.Driver;
import java.util.concurrent.TimeUnit;
import java.util.regex.Pattern;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.testng.Reporter;

import pageObjects.BaseClass;
import utility.Constant;
import utility.ExcelUtils;
import utility.Extend_Report;
import utility.Reusable;
import utility.Utils;
     
   
    public class ReviewQuote {
    	
    	  private static XSSFSheet ExcelWSheet;
          private static XSSFWorkbook ExcelWBook;
          private static XSSFCell Cell;
          private static XSSFRow Row;
         
         
        public  static void Execute(int iTestCaseRow,String TC_Name) throws Exception{
        	WebDriver driver=BaseClass.driver;
        	String class_name1=Thread.currentThread().getStackTrace()[1].getClassName();
        	String[] parts1 = class_name1.split(Pattern.quote("."));
        	String class_name=parts1[1];
        	
        	try{      		
        	     
        		FileInputStream ExcelFile = new FileInputStream(Constant.Path_Excel + Constant.File_TestData);
            	//System.out.println(Constant.Path_Excel + Constant.File_TestData);
                // Access the required test data sheet
        		
                ExcelWBook = new XSSFWorkbook(ExcelFile);
                ExcelWSheet = ExcelWBook.getSheet(class_name);
            	int noOfColumns = ExcelWSheet.getRow(0).getPhysicalNumberOfCells();                	
            	
				for (int i=1;i < noOfColumns;i++)
            	{
            	
					String  val=ExcelWSheet.getRow(0).getCell(i).getStringCellValue();
					
					int Valpos = val.indexOf("_");
					
					if (Valpos > 0)
					{
						
						String obj=val;
						String[] parts = val.split(Pattern.quote("_"));
						
						String Class_Name = parts[0];
						
						String Object_Type = parts[1];
						
						if (Class_Name.equalsIgnoreCase(class_name))
						{
							 
							
							switch (Object_Type){
							case "OpenBrowser":
								
								driver=Reusable.OpenBrowser(driver,iTestCaseRow,obj, Class_Name,TC_Name);
								  
								break;
							
							case "AutoitUpload":
								
								Reusable.AutoITUpload(driver,iTestCaseRow,obj, Class_Name,TC_Name);
								  
								break;
							case "FrameChange":
								
								Reusable.FrameChange(driver,iTestCaseRow,obj, Class_Name,TC_Name);
								  
								break;
							case "AysnWait":
								
								Reusable.AysnWait(driver,iTestCaseRow,obj, Class_Name,TC_Name);
								  
								break;
							
							case "Edit":
								
								Reusable.Enter_Value_WebElement(driver,iTestCaseRow,obj, Class_Name,TC_Name);								
								  
								break;
								
							case "EditWait":
								Thread.sleep(2000);
								Reusable.Enter_Value_WebElement(driver,iTestCaseRow,obj, Class_Name,TC_Name);
																
								break;
							case "MouseOver":
								
								Reusable.MouseOver_WebElement(driver,iTestCaseRow,obj, Class_Name,TC_Name);
															
								break;	
							case "WaitFor":
								
								Reusable.WaitForElement(driver,iTestCaseRow,obj, Class_Name,TC_Name);
																
								break;
							case "MouseOverElementdblClick":
								
								Reusable.MouseOver_WebElementdblClick(driver,iTestCaseRow,obj, Class_Name,TC_Name);
																
								break;
							case "MouseOverElementClick":
								
								Reusable.MouseOver_WebElementClick(driver,iTestCaseRow,obj, Class_Name,TC_Name);
																
								break;
							case "Button":
								
								Reusable.Click_Button(driver,iTestCaseRow, obj, Class_Name,TC_Name);
								
								
								break;
							case "ButtonOptional":
								
								Reusable.Click_Button_Optional(driver,iTestCaseRow, obj, Class_Name,TC_Name);
								
								
								break;
							case "TypeList":
								Reusable.Select_Value_DropDown(driver,iTestCaseRow, obj, Class_Name,TC_Name);
							
								
								break;
								
							case "TypeListValue":
								Reusable.Select_Value_DropDownValue(driver,iTestCaseRow, obj, Class_Name,TC_Name);
							
								
								break;	
							case "CheckWebElement":
								Reusable.Check_WebElement(driver,iTestCaseRow, obj, Class_Name,TC_Name);
								
								
								break;
							case "Checkbox":
								Reusable.Click_CheckBox(driver,iTestCaseRow, obj, Class_Name,TC_Name);
								
								
								break;
							case "Radio":
								Reusable.Click_RadioButton(driver,iTestCaseRow, obj, Class_Name,TC_Name);
							
								
								break;				
							case "ErrorValidation":
								Reusable.ErrorValidation(driver,iTestCaseRow, obj, Class_Name,TC_Name);
							
								
								break;	
							case "Browserclose":
								Reusable.closeBrowser(driver,iTestCaseRow, obj, Class_Name,TC_Name);
								  
								break;
								
								
							default:
								
								break;
							        		
							}
			        		 
			        		
			        		  
							
						}
							
						
					}
				
					
            	}
        		
				     	               
                
                Reporter.log("SignIn Action is successfully perfomed");
                
        	}
        	
        	catch(Exception e)
        	{
        		
        		System.out.println("Error in ReviewQuote :"+e.toString());
        		Extend_Report.AddReport("Error in ReviewQuote Page ", "Fail");
       		 
       		 	Extend_Report.AddReport(e.toString(), "Fail");
       		 
        		ExcelUtils.setExcelFile(Constant.Path_Excel+Constant.File_RunManager, "Run_Manager");		
       		 
        		int val_run=ExcelUtils.getRowContains_manager(TC_Name, 0, "Run_Manager");
       		 
       		 	ExcelUtils.setCellData_Manager("FAIL", val_run, 2);
        		
        		
        	}
        	
        }
    }
