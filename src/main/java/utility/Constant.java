package utility;

import org.openqa.selenium.By;

public class Constant {
	 	public static String currentDir = System.getProperty("user.dir");
	 	
	    public static final String DEV_URL = "https://audanet-cae.audatex.com.au/sso/login?service=https%3a%2f%2faudanet-cae.audatex.com.au%2fbre";
	    public static final String SYS_URL = "https://audanet-cae.audatex.com.au/sso/login?service=https%3a%2f%2faudanet-cae.audatex.com.au%2fbre";
	    public static final String UAT_URL = "https://audanet-cae.audatex.com.au/sso/login?service=https%3a%2f%2faudanet-cae.audatex.com.au%2fbre";
	    public static final String Username = "qbeSelvakumar";
	    public static final String Password ="EwUuQhp4uC"; 
		public static final String Path_Excel = currentDir+"//src//main//java//testData//";
		public static final String Sikuli_Path = currentDir+"//src//main//java//Sikuli//";
		public static final String AutoIT_UploadPath = currentDir+"//src//main//java//Upload//";
		public static final String File_TestData = "TestData.xlsx";
		public static final String File_RunManager = "Run_Manager.xlsx";	
		public static final String Path_ScreenShot = currentDir+"//src//main//java//Screenshots//";
		public static final String Path_IEDriver = currentDir+"//src////main//java//Jar//IEDriverServer.exe";
		public static final String Path_FireFoxDriver = currentDir+"//src//main//java//Jar//geckodriver.exe";
		public static final String Path_ChromeDriver = currentDir+"//src//main//java//Jar//chromedriver.exe";
		public static final String Path_EDGEDriver = currentDir+"//src//main//java//Jar//MicrosoftWebDriver.exe";
		public static  String Outlook_EmailID="selvakumar.c@mphasis.com";
		public static  String Outlook_EmailPwd="Selva@996";
		public static  String TC_Name;
		public static  String Global_ClaimNO;
		public static  String Global_BrowserType;
		public static  By LocatorV;
		public static  String TC_Status="";
	}
