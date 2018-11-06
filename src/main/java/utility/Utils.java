
package utility;
import java.io.File;
import java.sql.Array;
import java.util.concurrent.TimeUnit;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.support.ui.WebDriverWait;

import pageObjects.BaseClass;

public class Utils {
		public static WebDriver driver = null;
		public static String ENV_URL;
	public static WebDriver OpenBrowser(int iTestCaseRow) throws Exception{
		String sBrowserName;		
		int BrowserCOl = ExcelUtils.get_column("Login_Browser_Type","Test_Case");	
		
		 int BrowserENV = ExcelUtils.get_column("Login_ENV","Test_Case");
		
		 
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
		sBrowserName = ExcelUtils.getCellData(iTestCaseRow, BrowserCOl);	
		if(sBrowserName.equals("IE")){
			//Utils.takeScreenshot(driver, "Browser_Invoke");
			//System.setProperty("webdriver.gecko.driver", "C:\\Users\\selvakumar.c\\Desktop\\Selenium\\jar\\geckodriver-v0.18.0-win64\\geckodriver.exe");
			//driver = new FirefoxDriver();
			System.out.println("Internet Explorer is selected");
			System.setProperty("webdriver.ie.driver",Constant.Path_IEDriver);
			driver = new InternetExplorerDriver();
			Log.info("New driver instantiated");
		    driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
		    Log.info("Implicit wait applied on the driver for 10 seconds");
		    //driver.get("https://www.google.com.au/webhp?ie=utf-8&oe=utf-8&gws_rd=cr&ei=VfZvWYvnAoG80ATDsaTQAQ");
		   
		    driver.get(ENV_URL);
		    Extend_Report.AddReport("URL Launched "+ENV_URL , "Pass");
		    Log.info("Web application launched successfully");
			}
		
		else if (sBrowserName.equals("Mozilla")){
			//Utils.takeScreenshot(driver, "Browser_Invoke");
			System.setProperty("webdriver.gecko.driver", Constant.Path_FireFoxDriver);
			
			System.out.println("Mozilla  Explorer is selected");			
			driver = new FirefoxDriver();
			Log.info("New driver instantiated");
		    driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
		    Log.info("Implicit wait applied on the driver for 10 seconds");
		    driver. manage(). window().maximize();
		    driver.get(ENV_URL);
		    Extend_Report.AddReport("URL Launched "+ENV_URL , "Pass");
		    Log.info("Web application launched successfully");
			
		}
		else if (sBrowserName.equals("Chrome")){
			//Utils.takeScreenshot(driver, "Browser_Invoke");
			System.setProperty("webdriver.chrome.driver", Constant.Path_ChromeDriver);
			driver = new ChromeDriver();
			System.out.println("Chrome Browser is selected");	
			
			Log.info("New driver instantiated");
		    driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
		    Log.info("Implicit wait applied on the driver for 10 seconds");
		    driver. manage(). window().maximize();
		    driver.get(ENV_URL);
		    Extend_Report.AddReport("URL Launched "+ENV_URL , "Pass");
		    Log.info("Web application launched successfully");
			
		}
		}catch (Exception e){
			//Utils.takeScreenshot(driver, "Browser_Invoke");
			System.out.println("Browser Error");
			Log.error("Class Utils | Method OpenBrowser | Exception desc : "+e.getMessage());
		}
		new BaseClass(driver);	
		return driver;
	}
	
	public static String getTestCaseName(String sTestCase)throws Exception{
		String value = sTestCase;
		try{
			int posi = value.indexOf("@");
			value = value.substring(0, posi);
			posi = value.lastIndexOf(".");	
			value = value.substring(posi + 1);
			return value;
				}catch (Exception e){
			Log.error("Class Utils | Method getTestCaseName | Exception desc : "+e.getMessage());
			throw (e);
					}
			}
	
	
	
	
	 
	 
	}
