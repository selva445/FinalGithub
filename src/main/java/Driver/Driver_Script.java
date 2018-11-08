package Driver;

import java.io.File;
import java.lang.reflect.Method;
import java.util.ArrayList;

import org.apache.log4j.xml.DOMConfigurator;
import org.openqa.selenium.WebDriver;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;

import pageObjects.BaseClass;
import utility.Constant;
import utility.ExcelUtils;
import utility.Extend_Report;
import utility.Log;
import utility.Utils;

public class Driver_Script {
	public WebDriver driver;
	private String sTestCaseName;
	private int iTestCaseRow;	
	
	ArrayList<String> arr = new ArrayList<String>();
	
	
  @BeforeMethod
  public void beforeMethod() throws Exception {
	  
	  
	  
	  	DOMConfigurator.configure("log4j.xml");	 
	  	sTestCaseName = this.toString();
	  	sTestCaseName = Utils.getTestCaseName(this.toString());	  	
	  	ExcelUtils.GetTCExecute(arr);	  
	  	//
	  	
        }
  
  @SuppressWarnings({ "unused", "rawtypes" })
  
@Test
  public void f() throws Exception {
	  
	  for(String  z:arr)
		 {
			  ArrayList<String> keywords1 = new ArrayList<String>();
			 
			 try{
				 Class noparams[] = {};
				
				 int ntestcase = ExcelUtils.get_column("TestCaseName","Test_Case");			
				 iTestCaseRow = ExcelUtils.getRowContains(z,ntestcase,"Test_Case");						 
				  
				  ExcelUtils.GetKeywords(keywords1,iTestCaseRow);			  
				  for(String  k:keywords1)
				  {
					  
					  	//split string based on _
					  	boolean k2=k.contains("_");
					  	if (k2)
					  	{
					  		
					  		String page_name = k.split("_")[0];
		             		String page_iterator = k.split("_")[1];	
		             		ExcelUtils.create_ClassFiles(page_name);
					  		
					  			
					  	}
					  	else
					  	{
					  		ExcelUtils.create_ClassFiles(k);
					  		
					  	}
					  	
					  
				  }
				  
			
				 
			  }catch (Exception e){
				  System.out.println("Driver Script Error " +e.toString());
				  
				 
		 }
	}
	  

	  
	  
	  for(String  i:arr)
	 {
		  Constant.TC_Status="";
		  ArrayList<String> keywords = new ArrayList<String>();
		 
		 try{
			 Class noparams[] = {};
			 Extend_Report.Set_Header(i); 
			 int ntestcase = ExcelUtils.get_column("TestCaseName","Test_Case");			
			 iTestCaseRow = ExcelUtils.getRowContains(i,ntestcase,"Test_Case");	
			 String Browsername = ExcelUtils.getCellData(iTestCaseRow, 3);
			 // driver = Utils.OpenBrowser(iTestCaseRow);		
			//  new BaseClass(driver);		 			 
			  Log.startTestCase(i); 
			  ExcelUtils.GetKeywords(keywords,iTestCaseRow);			  
			  for(String  k:keywords)
			  {
				  
				  	//split string based on _
				  	boolean k2=k.contains("_");
				  	if (k2)
				  	{
				  		int val_run=ExcelUtils.getRowContains_manager(i, 0, "Run_Manager");
						 
						 String status=ExcelUtils.getCellData(val_run, 2);
						 
						 if (!(status.equalsIgnoreCase("")))
						 {
							 ExcelUtils.setCellData_Manager("", val_run, 2);
						 }
				  		String page_name = k.split("_")[0];
	             		String page_iterator = k.split("_")[1];	
	             		//ExcelUtils.create_ClassFiles(page_name);
				  		System.out.println(" Executing Component Name :  " + page_name);
				  		System.out.println(" Executing Component Iteration :  " + page_iterator);
					  	Extend_Report.AddReport("Calling Screen  :"+ page_name, "info"); 
					  	Extend_Report.AddReport("Calling Screen Iteration :"+ page_iterator, "info"); 
					  	String k1="appModules"+"."+page_name;
					  	String dogClassName = k1;
				        Class<?> dogClass = Class.forName(dogClassName); // convert string classname to class
				        Object dog = dogClass.newInstance(); // invoke empty constructor
				        String methodName = "";
				        methodName = "Execute";
				        Method setNameMethod = dog.getClass().getMethod(methodName,int.class,String.class);
				        int rowval = ExcelUtils.getRowContains_iterator(i,0,page_name,page_iterator);
				        String Tcpath=Constant.currentDir;				        
						boolean success = (new File(Constant.Path_ScreenShot+i+Browsername)).mkdirs();				     
				        Constant.TC_Name=Constant.Path_ScreenShot+i+Browsername+"//";   
				      
				        setNameMethod.invoke(dog,rowval,i); // pass arg
				  			
				  	}
				  	else
				  	{
				  		int val_run=ExcelUtils.getRowContains_manager(i, 0, "Run_Manager");
						 
						 String status=ExcelUtils.getCellData(val_run, 2);
						 
						 if (!(status.equalsIgnoreCase("")))
						 {
							 ExcelUtils.setCellData_Manager("", val_run, 2);
						 }
				  		//ExcelUtils.create_ClassFiles(k);
				  		System.out.println(" Executing Component Name :  " + k);
					  	Extend_Report.AddReport("Calling Screen  :"+ k, "info"); 
					  	String k1="appModules"+"."+k;
					  	String dogClassName = k1;
				        Class<?> dogClass = Class.forName(dogClassName); // convert string classname to class
				        Object dog = dogClass.newInstance(); // invoke empty constructor
				        String methodName = "";
				        methodName = "Execute";
				        Method setNameMethod = dog.getClass().getMethod(methodName,int.class,String.class);
				        int rowval = ExcelUtils.getRowContains(i,0,k);	
				        String Tcpath=Constant.currentDir;
				       
				        boolean success = (new File(Constant.Path_ScreenShot+i+Browsername)).mkdirs();				     
				        Constant.TC_Name=Constant.Path_ScreenShot+i+Browsername+"//";
				       // System.out.println(Constant.TC_Name);
				        setNameMethod.invoke(dog,rowval,i); // pass arg
				   	 
				  	}
				  	
				  
			  }
			  
		
			  
			  Log.endTestCase(i);		
			  Thread.sleep(3000);
			  WebDriver driver=BaseClass.driver;
		  		if (driver!=null)
				  {
		  			driver.quit();
				  }	 
		    
		  	Extend_Report.Clean_Up();
			 
		  }catch (Exception e){
			  System.out.println("Driver Script Error " +e.toString());
			  Log.endTestCase(sTestCaseName);
			  Thread.sleep(3000);
			  WebDriver driver=BaseClass.driver;
		  		if (driver!=null)
				  {
		  			driver.quit();
				  }	 
		  		 ExcelUtils.setExcelFile(Constant.Path_Excel+Constant.File_RunManager, "Run_Manager");		
				 
				 int val_run=ExcelUtils.getRowContains_manager(i, 0, "Run_Manager");
				 
				 ExcelUtils.setCellData_Manager("FAIL", val_run, 2);
		  		Extend_Report.Clean_Up();  
			  Log.error(e.getMessage());
			 
		  }
			
		 ExcelUtils.setExcelFile(Constant.Path_Excel+Constant.File_RunManager, "Run_Manager");		
		 
		 int val_run=ExcelUtils.getRowContains_manager(i, 0, "Run_Manager");
		 
	/*	 String status=ExcelUtils.getCellData(val_run, 2);	
		
		 System.out.println("Final Status of TC "+ i +status);
		 
		 if (!(status.equalsIgnoreCase("FAIL")))
		 {
			 ExcelUtils.setCellData_Manager("PASS", val_run, 2);
		 }
		 else if(status.equalsIgnoreCase(""))
		 {
			 ExcelUtils.setCellData_Manager("PASS", val_run, 2); 
		 }*/
		 String status=Constant.TC_Status;	
		
		 System.out.println("Final Status of TC  "+ i +status);
		 
		 if (!(status.equalsIgnoreCase("FAIL")))
		 {
			 ExcelUtils.setCellData_Manager("PASS", val_run, 2);
		 }
		 else if(status.equalsIgnoreCase(""))
		 {
			 ExcelUtils.setCellData_Manager("PASS", val_run, 2); 
		 }
		 else
		 {
			 ExcelUtils.setCellData_Manager("FAIL", val_run, 2); 
		 }
	 }
	  
	  
	  
	  }
  
		
		
  @AfterMethod
  public void afterMethod() throws Exception {	   
	
	  //Log.endTestCase(sTestCaseName);
	  
	  		if (driver!=null)
		  {
	  			driver.quit();
		  }
	 
	    
	 //   Extend_Report.Clean_Up();
	 
	   
  		}
}
