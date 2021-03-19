package api_automation;

import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;
import org.openqa.selenium.WebDriver;

import API_Automation_THS.MH;
import API_Automation_THS.api_DP1;
import API_Automation_THS.api_DP3;
import API_Automation_THS.api_HO3;
import API_Automation_THS.api_HO4;
import API_Automation_THS.api_HO6;
import API_Automation_THS.api_MH;
import Applicability_Package.Constant;
import Applicability_Package.MH_ExcelUtils;
import API_Automation_THS.api_MH;

public class run_api_automation {
 private static  Logger log = Logger.getLogger(MH.class.getName()+" ----------------------------------");
    //given, When ,There
    
//  =================================================================================================================================================================================================================
   /*                  OVERVIEW
    *                 ==========
    * 
    *   Create a universal scripting to run api's based on policy form 
    *   
    *   Example :
    *              if data is provided with random policy fomrs (D1,D3,H3,H6,H4,MH) instead of running different classes based in policy form,
    *              this program should automatically read policy form and run accordinglly.
    *     
    * 
    */
//  =================================================================================================================================================================================================================        


 
 
private static WebDriver driver = null;

public static void main(String[] args) throws Exception {  
	String log4jConfPath = "C:\\Users\\vdaru\\eclipse-workspace\\API_Automation_THS\\src\\API_Automation_THS\\log4j.properties";
        PropertyConfigurator.configure(log4jConfPath);
        
  // read excel and check policy form and print name of the form     
        
      MH_ExcelUtils.setExcelFile(Constant.MH_Path_TestData + Constant.MH_File_TestData, "sheet1");		
	int i, j;
	    for (j = 0; j < MH_ExcelUtils.ExcelWSheet.getPhysicalNumberOfRows(); j++) {    			
	        for (int a = 3; a <  7; a++) {
		       log.info("READ POLICY FORM FORM DATA SHEET");
		       String Policy_Form =        MH_ExcelUtils.getCellData(a, 1);		      
		       log.info("POLICY FORM =====  "  + Policy_Form);
//		       System.out.println(Policy_Form);
		
		 switch (Policy_Form)  {
		 
		 case "MC":
		     log.info("policy form is MH please run script for MH");   
		     api_MH.Execute(driver);		     
		     break;
		 case "DP-1":
		     log.info("policy form is DP-1 please run script for DP-1");
		     api_DP1.Execute(driver);
		     break;
		 case "DP-3":
		     log.info("policy form is DP-3 please run script for DP-3");
		     api_DP3.Execute(driver);
		     break;
		 case "HO-3":
		     log.info("policy form is HO-3 please run script for HO-3");
		     api_HO3.Execute(driver);
		     break;
		 case "HO-6":
		     log.info("policy form is HO-6 please run script for HO-6");
		     api_HO6.Execute(driver);
		     break;
		 case "HO-4":
		     log.info("policy form is HO-4 please run script for HO-4");
		     api_HO4.Execute(driver);
		     break;
		 }

		 
	        }
	        
	    }
	}
    }

