package utility;
            import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.lang.reflect.Array;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.format.CellFormatType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

    public class ExcelUtils {
                private static XSSFSheet ExcelWSheet;
                private static XSSFWorkbook ExcelWBook;
                private static XSSFCell Cell;
                private static XSSFRow Row;
            //This method is to set the File path and to open the Excel file, Pass Excel Path and Sheetname as Arguments to this method
            public static void setExcelFile(String Path,String SheetName) throws Exception {
                   try {
                       // Open the Excel file
                    FileInputStream ExcelFile = new FileInputStream(Path);
                    // Access the required test data sheet
                    ExcelWBook = new XSSFWorkbook(ExcelFile);
                    ExcelWSheet = ExcelWBook.getSheet(SheetName);
                    Log.info("Excel sheet opened");
                    } catch (Exception e){
                        throw (e);
                    }
            }
            
            //to get column number
            
            public static int get_column(String colname, String sheet) throws Exception{
                try{
             	  
                	FileInputStream ExcelFile = new FileInputStream(Constant.Path_Excel + Constant.File_TestData);
                	//System.out.println(Constant.Path_Excel + Constant.File_TestData);
                    // Access the required test data sheet
                    ExcelWBook = new XSSFWorkbook(ExcelFile);
                 
                    ExcelWSheet = ExcelWBook.getSheet(sheet);                    
                    
                	int noOfColumns = ExcelWSheet.getRow(0).getPhysicalNumberOfCells();                	
                	
					for (int i=0;i <= noOfColumns;i++)
                	{
                	
						String  val=ExcelWSheet.getRow(0).getCell(i).getStringCellValue();
						
						
						
						if (val.equalsIgnoreCase(colname))
						{
							
							return i;
							
						}
                	}
                	
                   }catch (Exception e){
                	   
                     System.out.println("Excel Exception "+e.toString());
                   }
				return 0;
				
         }
            
            //This method is to read the test data from the Excel cell, in this we are passing parameters as Row num and Col num
            public static String getCellData(int RowNum, int ColNum) throws Exception{
                   try{
                	   Cell = ExcelWSheet.getRow(RowNum).getCell(ColNum);      	
                	   Cell.setCellType(Cell.CELL_TYPE_STRING);              
                	  
                      String CellData = Cell.getStringCellValue(); 
                       
                      return CellData;
                      }
                   
                   		catch (Exception e){
                        return"";
                      }
            }
            
            
            //This method is to write in the Excel cell, Row num and Col num are the parameters
            @SuppressWarnings("static-access")
			public static void setCellData(String Result,  int RowNum, int ColNum) throws Exception    {
                   try{
                      Row  = ExcelWSheet.getRow(RowNum);
                    Cell = Row.getCell(ColNum, Row.RETURN_BLANK_AS_NULL);
                    if (Cell == null) {
                        Cell = Row.createCell(ColNum);
                        Cell.setCellValue(Result);
                        } else {
                            Cell.setCellValue(Result);
                        }
          // Constant variables Test Data path and Test Data file name
                          FileOutputStream fileOut = new FileOutputStream(Constant.Path_Excel + Constant.File_TestData);
                          ExcelWBook.write(fileOut);
                          fileOut.flush();
                        fileOut.close();
                        }catch(Exception e){
                            throw (e);
                    }
                }
            public static void setCellData_Manager(String Result,  int RowNum, int ColNum) throws Exception    {
                try{
                   Row  = ExcelWSheet.getRow(RowNum);
                 Cell = Row.getCell(ColNum, Row.RETURN_BLANK_AS_NULL);
                 if (Cell == null) {
                     Cell = Row.createCell(ColNum);
                     Cell.setCellValue(Result);
                     } else {
                         Cell.setCellValue(Result);
                     }
       // Constant variables Test Data path and Test Data file name
                       FileOutputStream fileOut = new FileOutputStream(Constant.Path_Excel + Constant.File_RunManager);
                       ExcelWBook.write(fileOut);
                       fileOut.flush();
                     fileOut.close();
                     }catch(Exception e){
                         throw (e);
                 }
             }
        	public static int getRowContains(String sTestCaseName, int colNum,String sheet) throws Exception{
        		int i;
        		try {
        			
        			ExcelUtils.setExcelFile(Constant.Path_Excel + Constant.File_TestData,sheet);
        			
        			int rowCount = ExcelUtils.getRowUsed();
        			
        			for ( i=0 ; i<rowCount; i++){
        				if  (ExcelUtils.getCellData(i,colNum).equalsIgnoreCase(sTestCaseName)){
        					break;
        				}
        			}
        			return i;
        				}catch (Exception e){
        			Log.error("Class ExcelUtil | Method getRowContains | Exception desc : " + e.getMessage());
        			throw(e);
        			}
        		}
        	public static int getRowContains_iterator(String sTestCaseName, int colNum,String sheet,String iterationparam) throws Exception{
        		int i;
        		try {
        			
        			ExcelUtils.setExcelFile(Constant.Path_Excel + Constant.File_TestData,sheet);
        			
        			int rowCount = ExcelUtils.getRowUsed();
        			
        			for ( i=0 ; i<=rowCount; i++){
        			
        				
        			
        				if  ((ExcelUtils.getCellData(i,colNum).equalsIgnoreCase(sTestCaseName)) && (ExcelUtils.getCellData(i,1).equalsIgnoreCase(iterationparam)))
        						{
        					
        					break;
        				}
        			}
        			return i;
        				}catch (Exception e){
        			Log.error("Class ExcelUtil | Method getRowContains | Exception desc : " + e.getMessage());
        			throw(e);
        			}
        		}
        	public static int getRowContains_manager(String sTestCaseName, int colNum,String sheet) throws Exception{
        		int i;
        		try {
        			
        			ExcelUtils.setExcelFile(Constant.Path_Excel + Constant.File_RunManager,sheet);
        			
        			int rowCount = ExcelUtils.getRowUsed();
        			
        			for ( i=0 ; i<rowCount; i++){
        				if  (ExcelUtils.getCellData(i,colNum).equalsIgnoreCase(sTestCaseName)){
        					break;
        				}
        			}
        			return i;
        				}catch (Exception e){
        			Log.error("Class ExcelUtil | Method getRowContains_manager | Exception desc : " + e.getMessage());
        			throw(e);
        			}
        		}
        	public static void GetTCExecute(ArrayList<String> tc) throws Exception{
        		
        		try {
        			
        			 ExcelUtils.setExcelFile(Constant.Path_Excel + Constant.File_RunManager,"Run_Manager");
        			 int rowCount = ExcelUtils.getRowUsed();        			 
        		
         			for ( int i = 1 ; i<rowCount+1; i++){
         				
         				String  val=ExcelUtils.getCellData(i, 1);    				
         				
         				
         				if  (val.equalsIgnoreCase("Yes")){        					
         					      					
         					
         					tc.add(ExcelUtils.getCellData(i,0));    				
         					
         					//System.out.println(tc); 
         				}
         			}
         			
        			
        		}
        			catch (Exception e){
        			Log.error("Class ExcelUtil | Method GetTCExecute | Exception desc : " + e.getMessage());
        			System.out.println("Exception RunManager");  	
        			System.out.println(e.getMessage());        			
        			
        			}
				
        		}
        	
  //*****************************************************************************************************************      	
 public static void create_ClassFiles(String Excel_Class) throws Exception
        	{
        		String currentDir = System.getProperty("user.dir");
        		        		
        		String Spath = currentDir+"\\src\\main\\java\\appModules\\"+"Template"+".java";
        		File SourceFile= new File(Spath);   
        		
        		String Dpath = currentDir+"\\src\\main\\java\\appModules\\"+Excel_Class+".java";
        		File DestFile= new File(Dpath);
        		
        		try {
        			//   File file = new File(Dpath);
        	        if (!DestFile.exists()) {
        	        	
        	        		//DestFile.createNewFile();        	           		
        	           		FileUtils.copyFile(SourceFile, DestFile);
        	           		ExcelUtils.Replace_Files(Dpath, "Template", Excel_Class);
        	           		//System.out.println("File  Created "+ Excel_Class + " As it did not Exist");
        	           		//IProject[] pt=ResourcesPlugin.getWorkspace().getRoot().getProjects();
        	           		
        	           		//for(IProject pro : pt)
        	           	//	{
        	           	//		pro.refreshLocal(IResource.DEPTH_INFINITE,new NullProgressMonitor());
        	           			
        	           	//	}
        	           		
        	        }      	           		
        	        else
        	        {
        	        	//System.out.println("File Not Created "+ Excel_Class + " As it already Exists");
        	        }
        	        
        	        
        	    } catch (IOException e) {
        	    	System.out.println(e.getMessage());
        	    }
        		
  }
        	  //*****************************************************************************************************************    
  public static void Replace_Files(String filePath, String oldString, String newString) throws Exception   
  
  {
	  
	  File fileToBeModified = new File(filePath);
      
      String oldContent = "";
       
      BufferedReader reader = null;
       
      FileWriter writer = null;
       
      try
      {
          reader = new BufferedReader(new FileReader(fileToBeModified));
           
          //Reading all the lines of input text file into oldContent
           
          String line = reader.readLine();
           
          while (line != null) 
          {
              oldContent = oldContent + line + System.lineSeparator();
               
              line = reader.readLine();
          }
           
          //Replacing oldString with newString in the oldContent
           
          String newContent = oldContent.replaceAll(oldString, newString);
           
          //Rewriting the input text file with newContent
           
          writer = new FileWriter(fileToBeModified);
           
          writer.write(newContent);
      }
      catch (IOException e)
      {
          e.printStackTrace();
      }
      finally
      {
          try
          {
              //Closing the resources
               
              reader.close();
               
              writer.close();
          } 
          catch (IOException e) 
          {
              e.printStackTrace();
          }
      }
	  
  }
        	
        	
	public static void GetKeywords(ArrayList<String> keywords,int rows) throws Exception{
        		
        		try {
        			
        			 ExcelUtils.setExcelFile(Constant.Path_Excel + Constant.File_TestData,"Test_Case");
        			 
        			 int colmax=ExcelWSheet.getRow(rows).getLastCellNum(); 
        			
         			for ( int i = 4 ; i < colmax; i++){
         				
         				String  val=ExcelUtils.getCellData(rows, i);    				
         				
         				
         				if(!(val.trim().isEmpty())){        					
         					      					
         					
         					keywords.add(val);    				
         					
         					
         				}
         			}
         			
        			
        		}
        			catch (Exception e){
        			Log.error("Class GetKeywords | Method GetKeywords | Exception desc : " + e.getMessage());
        			System.out.println("Exception GetKeywords");  	
        			System.out.println(e.getMessage());        			
        			
        			}
				
        		}
        	public static int getRowUsed() throws Exception {
        		try{
        			int RowCount = ExcelWSheet.getLastRowNum();
        			Log.info("Total number of Row used return as < " + RowCount + " >.");		
        			return RowCount;
        		}catch (Exception e){
        			Log.error("Class ExcelUtil | Method getRowUsed | Exception desc : "+e.getMessage());
        			System.out.println(e.getMessage());
        			throw (e);
        		}
        		
        		
        	}
    }
    
    