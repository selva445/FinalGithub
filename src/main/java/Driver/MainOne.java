/*    */ package Driver;
/*    */ 
/*    */ import java.io.PrintStream;
/*    */ import java.text.DateFormat;
import java.text.ParseException;
/*    */ import java.text.SimpleDateFormat;
/*    */ import java.util.Calendar;
/*    */ import java.util.Date;

/*    */ import org.testng.TestListenerAdapter;
/*    */ import org.testng.TestNG;
/*    */ 
/*    */ 
/*    */ 
/*    */ 
/*    */ public class MainOne
/*    */ {
			   static Calendar cal = Calendar.getInstance();
	           DateFormat DF = new SimpleDateFormat("dd/MM/yyyy");
	           
	          
	           public static void main(String[] args) throws ParseException
	           {
				 Date date = new Date();
				// Date date1 = date.after(when)
			     SimpleDateFormat dt1 = new SimpleDateFormat("dd/MM/yyyy");
			     Date date1 = dt1.parse("13/10/2018");
			     
      
/* 20 */     Date TodayDate = cal.getTime();
     
/* 24 */     if (TodayDate.after(date1))
/*    */     {
/* 26 */       System.out.println("JAR REACHED THE EXPIRATION DATE");
/*    */     }
/*    */     else
/*    */     {
/* 30 */       TestListenerAdapter tla = new TestListenerAdapter();
/* 31 */       TestNG testng = new TestNG();
/* 32 */       Class[] classes = { Driver_Script.class };
/*    */       
/* 34 */       testng.setTestClasses(classes);
/* 35 */       testng.addListener(tla);
/* 36 */       testng.run();
/*    */     }
/*    */   }
/*    */ }


/* Location:              C:\Users\889128\Desktop\M-EzScriptor WS.jar!\Driver\MainOne.class
 * Java compiler version: 8 (52.0)
 * JD-Core Version:       0.7.1
 */