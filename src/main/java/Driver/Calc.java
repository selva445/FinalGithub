package Driver;

import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.safari.SafariDriver;
import org.sikuli.script.FindFailed;
import org.sikuli.script.Pattern;
import org.sikuli.script.Screen;

import utility.Constant;

public class Calc {

	public static void main(String[] args) throws FindFailed, InterruptedException {
		
		String val="A new Request for Quote Case with Assessment Number QBE00001577 is now created.";
		System.out.println(val);
		String val1=val.substring(52, 63);
		System.out.println(val1);
		 
		

	}

}
