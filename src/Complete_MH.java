
import io.restassured.RestAssured;
import io.restassured.path.xml.XmlPath;
import io.restassured.response.Response;
import static io.restassured.RestAssured.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Date;
import java.text.SimpleDateFormat;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import Applicability_Package.Constant;
import Applicability_Package.MH_ExcelUtils;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;


public class Complete_MH  extends MH_ExcelUtils{

    //given, When ,There
    
    public static void main(String[] args) throws IOException {
	
	    FileInputStream fis = new FileInputStream("C:\\Users\\vdaru\\Desktop\\API_Automation\\MH\\API_MH.xlsx");	 
	    XSSFWorkbook wb=new XSSFWorkbook(fis);
            XSSFSheet sheet=wb.getSheet("Sheet1");	
             
            int noofRows = sheet.getLastRowNum();
		System.out.println("the total number of Rows are " + "------ " + noofRows);
		String[] Headers1 = new String[noofRows];
		int i, j;
		for (j = 0; j < noofRows; j++) {    			
			 for (int a = 3; a < noofRows; a++) {	
			     System.out.println("im in Row   :" + a);

	     String API_BETA= Constant.URL_API_UAT;
	      
	     
	        XSSFRow rowAPIKey=sheet.getRow(a);			                                              			
	        XSSFCell CollAPIKey=rowAPIKey.getCell(1);                       			                   			                                                       
	        String APIKey_UAT=CollAPIKey.getStringCellValue();	         
	      
	        long dAOP = (long) wb.getSheetAt(0).getRow(a).getCell(2).getNumericCellValue();
	        String AOP = String.valueOf(dAOP); 
	        
	        long dCovA = (long) wb.getSheetAt(0).getRow(a).getCell(3).getNumericCellValue();
	        String CovA = String.valueOf(dCovA); 
	        
	        long dCovB = (long) wb.getSheetAt(0).getRow(a).getCell(4).getNumericCellValue();
	        String CovB = String.valueOf(dCovB); 
	        
	        long dCovC = (long) wb.getSheetAt(0).getRow(a).getCell(5).getNumericCellValue();
	        String CovC = String.valueOf(dCovC); 
	        
	        long dCovE = (long) wb.getSheetAt(0).getRow(a).getCell(6).getNumericCellValue();
	        String CovE = String.valueOf(dCovE); 
	        
	        long dCovF = (long) wb.getSheetAt(0).getRow(a).getCell(7).getNumericCellValue();
	        String CovF = String.valueOf(dCovF); 
	        
	        XSSFRow rowNamedStormDed=sheet.getRow(a);			                                              			
	        XSSFCell CollNamedStormDed=rowNamedStormDed.getCell(9);                       			                   			                                                       
	        String NamedStormDed=CollNamedStormDed.getStringCellValue();
	        
	        long dNamedStormDed = (long) wb.getSheetAt(0).getRow(a).getCell(8).getNumericCellValue();
	        String Hurrican_ded = String.valueOf(dNamedStormDed);
	        
	        long dWater_Ded = (long) wb.getSheetAt(0).getRow(a).getCell(10).getNumericCellValue();
	        String Water_Ded = String.valueOf(dWater_Ded); 
	        
	        long dWindHailDed = (long) wb.getSheetAt(0).getRow(a).getCell(11).getNumericCellValue();
	        String WindHailDed = String.valueOf(dWindHailDed); 
	       
	        XSSFRow rowEffectiveDate=sheet.getRow(a);			                                              			
	        XSSFCell CollEffectiveDate=rowEffectiveDate.getCell(12);                       			                   			                                                       
	        String EffectiveDate=CollEffectiveDate.getStringCellValue();
	        
		XSSFRow rowAssociationType=sheet.getRow(a);			                                              			
	        XSSFCell CollssociationType=rowAssociationType.getCell(13);                       			                   			                                                       
	        String AssociationType=CollssociationType.getStringCellValue();
	        
	        XSSFRow rowPaperLess=sheet.getRow(a);			                                              			
	        XSSFCell CollPaperLess=rowPaperLess.getCell(14);                       			                   			                                                       
	        String PaperLess=CollPaperLess.getStringCellValue();
	        
	        XSSFRow rowPayInFull=sheet.getRow(a);			                                              			
	        XSSFCell CollPayInFull=rowPayInFull.getCell(15);                       			                   			                                                       
	        String PayInFull=CollPayInFull.getStringCellValue();
	        
	        XSSFRow rowWaterDamage=sheet.getRow(a);			                                              			
	        XSSFCell CollWaterDamage=rowWaterDamage.getCell(16);                       			                   			                                                       
	        String WaterDamage=CollWaterDamage.getStringCellValue();
	        
	        XSSFRow rowAA=sheet.getRow(a);			                                              			
		XSSFCell CollAA=rowAA.getCell(18);                       			                   			                                                       
		String Address=CollAA.getStringCellValue();
	        
	        long dRisk_Street_number = (long) wb.getSheetAt(0).getRow(a).getCell(17).getNumericCellValue();
	        String Risk_Street_number = String.valueOf(dRisk_Street_number);
	        
	        XSSFRow rowRisk_Street_name=sheet.getRow(a);			                                              			
	        XSSFCell CollRisk_Street_name=rowRisk_Street_name.getCell(18);                       			                   			                                                       
	        String Risk_Street_name=CollRisk_Street_name.getStringCellValue();
	        
	        XSSFRow rowRisk_City=sheet.getRow(a);			                                              			
	        XSSFCell CollRisk_City=rowRisk_City.getCell(19);                       			                   			                                                       
	        String Risk_City=CollRisk_City.getStringCellValue();
	        
	        XSSFRow rowRisk_State=sheet.getRow(a);			                                              			
	        XSSFCell CollRisk_State=rowRisk_State.getCell(20);                       			                   			                                                       
	        String Risk_State=CollRisk_State.getStringCellValue();
	        
	        long dRisk_ZpiCode = (long) wb.getSheetAt(0).getRow(a).getCell(21).getNumericCellValue();
	        String Risk_ZpiCode = String.valueOf(dRisk_ZpiCode);
	        
	        long dConstructionYear = (long) wb.getSheetAt(0).getRow(a).getCell(22).getNumericCellValue();
	        String ConstructionYear = String.valueOf(dConstructionYear);
	       	        
	        XSSFRow rowDwellingLossSettlement=sheet.getRow(a);			                                              			
	        XSSFCell CollDwellingLossSettlement=rowDwellingLossSettlement.getCell(23);                       			                   			                                                       
	        String DwellingLossSettlement=CollDwellingLossSettlement.getStringCellValue();
	        
	        XSSFRow rowMH_AdvantageHome=sheet.getRow(a);			                                              			
	        XSSFCell CollMH_AdvantageHome=rowMH_AdvantageHome.getCell(24);                       			                   			                                                       
	        String MH_AdvantageHome=CollMH_AdvantageHome.getStringCellValue();
	        
	        XSSFRow rowMobileHomeType=sheet.getRow(a);			                                              			
	        XSSFCell CollMobileHomeType=rowMobileHomeType.getCell(25);                       			                   			                                                       
	        String MobileHomeType=CollMobileHomeType.getStringCellValue();
	        
	        long dMonthsUnOccupied = (long) wb.getSheetAt(0).getRow(a).getCell(26).getNumericCellValue();
	        String MonthsUnOccupied = String.valueOf(dMonthsUnOccupied);
	      
	        XSSFRow rowNear_Hyderate=sheet.getRow(a);			                                              			
	        XSSFCell CollNear_Hyderate=rowNear_Hyderate.getCell(27);                       			                   			                                                       
	        String Near_Hyderate=CollNear_Hyderate.getStringCellValue();
	        
	        XSSFRow rowOccupancy=sheet.getRow(a);			                                              			
	        XSSFCell CollOccupancy=rowOccupancy.getCell(28);                       			                   			                                                       
	        String Occupancy=CollOccupancy.getStringCellValue();
	        
	        XSSFRow rowOccupiedBY=sheet.getRow(a);			                                              			
	        XSSFCell CollOccupiedBY=rowOccupiedBY.getCell(29);                       			                   			                                                       
	        String OccupiedBY=CollOccupiedBY.getStringCellValue();
	        
	        XSSFRow rowParkStatus=sheet.getRow(a);			                                              			
	        XSSFCell CollParkStatus=rowParkStatus.getCell(30);                       			                   			                                                       
	        String ParkStatus=CollParkStatus.getStringCellValue();
	        
	        long dRooFYear = (long) wb.getSheetAt(0).getRow(a).getCell(31).getNumericCellValue();
	        String RooFYear = String.valueOf(dRooFYear);
	        
	        XSSFRow rowShortTermRentalSurecharge=sheet.getRow(a);			                                              			
	        XSSFCell CollShortTermRentalSurecharge=rowShortTermRentalSurecharge.getCell(32);                       			                   			                                                       
	        String ShortTermRentalSurecharge=CollShortTermRentalSurecharge.getStringCellValue();
	        
	        long dProtectionClass = (long) wb.getSheetAt(0).getRow(a).getCell(33).getNumericCellValue();
	        String ProtectionClass = String.valueOf(dProtectionClass);
	        
	        XSSFRow rowPersonalPropertyReplacementCost=sheet.getRow(a);			                                              			
	        XSSFCell CollPersonalPropertyReplacementCost=rowPersonalPropertyReplacementCost.getCell(34);                       			                   			                                                       
	        String PersonalPropertyReplacementCost=CollPersonalPropertyReplacementCost.getStringCellValue();
	        
	        XSSFRow rowEmail=sheet.getRow(a);			                                              			
	        XSSFCell CollEmail=rowEmail.getCell(35);                       			                   			                                                       
	        String Email=CollEmail.getStringCellValue();
	        
	        XSSFRow rowInsurenceScore=sheet.getRow(a);			                                              			
	        XSSFCell CollInsurenceScore=rowInsurenceScore.getCell(36);                       			                   			                                                       
	        String InsurenceScore=CollInsurenceScore.getStringCellValue();
	        
	        XSSFRow rowMailing_County=sheet.getRow(a);			                                              			
	        XSSFCell CollMailing_County=rowMailing_County.getCell(37);                       			                   			                                                       
	        String Mailing_County=CollMailing_County.getStringCellValue();
	        
	        long dMailing_Street_number = (long) wb.getSheetAt(0).getRow(a).getCell(38).getNumericCellValue();
	        String Mailing_Street_number = String.valueOf(dMailing_Street_number);
	        
	        XSSFRow rowMailing_Street_name=sheet.getRow(a);			                                              			
	        XSSFCell CollMailing_Street_name=rowMailing_Street_name.getCell(39);                       			                   			                                                       
	        String Mailing_Street=CollMailing_Street_name.getStringCellValue();
	        
	        XSSFRow rowMailingk_City=sheet.getRow(a);			                                              			
	        XSSFCell CollMailing_City=rowMailingk_City.getCell(40);                       			                   			                                                       
	        String Mailing_City=CollMailing_City.getStringCellValue();
	        
	        XSSFRow rowMailing_State=sheet.getRow(a);			                                              			
	        XSSFCell CollMailing_State=rowMailing_State.getCell(41);                       			                   			                                                       
	        String Mailing_State=CollMailing_State.getStringCellValue();
	        
	        long dMailingZipcode_ZpiCode = (long) wb.getSheetAt(0).getRow(a).getCell(42).getNumericCellValue();
	        String MailingZipcode_ZpiCode = String.valueOf(dMailingZipcode_ZpiCode);
	        
	        XSSFRow rowFirstName=sheet.getRow(a);			                                              			
	        XSSFCell CollFirstName=rowFirstName.getCell(43);                       			                   			                                                       
	        String FirstName=CollFirstName.getStringCellValue();
	        
	        XSSFRow rowLastName=sheet.getRow(a);			                                              			
	        XSSFCell CollLastName=rowLastName.getCell(44);                       			                   			                                                       
	        String LastName=CollLastName.getStringCellValue();
	       

	        XSSFRow rowDate_OF_Birth=sheet.getRow(a);			                                              			
	        XSSFCell CollDate_OF_Birth=rowDate_OF_Birth.getCell(45);                       			                   			                                                       
	        String Date_OF_Birth=CollDate_OF_Birth.getStringCellValue();
	           
		XSSFRow rowPrevious_County=sheet.getRow(a);			                                              			
		XSSFCell CollPrevious_County=rowPrevious_County.getCell(46);                       			                   			                                                       
		String Previous_County=CollPrevious_County.getStringCellValue();
		        
		long dPrevious_Street_number = (long) wb.getSheetAt(0).getRow(a).getCell(47).getNumericCellValue();
		String Previous_Street_number = String.valueOf(dPrevious_Street_number);
		        
	        XSSFRow rowPrevious_Street_name=sheet.getRow(a);			                                              			
		XSSFCell CollPrevious_Street_name=rowPrevious_Street_name.getCell(48);                       			                   			                                                       
		String Previous_Street=CollPrevious_Street_name.getStringCellValue();
		        
		XSSFRow rowPreviousk_City=sheet.getRow(a);			                                              			
		XSSFCell CollPrevious_City=rowPreviousk_City.getCell(49);                       			                   			                                                       
		String Previous_City=CollPrevious_City.getStringCellValue();
		        
		XSSFRow rowPrevious_State=sheet.getRow(a);			                                              			
		XSSFCell CollPrevious_State=rowPrevious_State.getCell(50);                       			                   			                                                       
		String Previous_State=CollPrevious_State.getStringCellValue();
		        
	        long dPreviousZipcode_ZpiCode = (long) wb.getSheetAt(0).getRow(a).getCell(51).getNumericCellValue();
		String PreviousZipcode_ZpiCode = String.valueOf(dPreviousZipcode_ZpiCode);
	     
		XSSFRow rowrNew_Purchase=sheet.getRow(a);			                                              			
		XSSFCell CollNew_Purchase=rowrNew_Purchase.getCell(52);                       			                   			                                                       
		String New_Purchase=CollNew_Purchase.getStringCellValue();
		
		XSSFRow rowOccasioanlRental_Surecharge=sheet.getRow(a);			                                              			
		XSSFCell CollOccasioanlRental_Surecharge=rowOccasioanlRental_Surecharge.getCell(53);                       			                   			                                                       
		String OccasioanlRental_Surecharge=CollOccasioanlRental_Surecharge.getStringCellValue();
		
		XSSFRow rowPrior_exp_Date=sheet.getRow(a);			                                              			
		XSSFCell CollPrior_exp_Date=rowPrior_exp_Date.getCell(54);                       			                   			                                                       
		String Prior_exp_Date=CollPrior_exp_Date.getStringCellValue();
		
         	long dSquareFootage = (long) wb.getSheetAt(0).getRow(a).getCell(55).getNumericCellValue();
		String SquareFootage = String.valueOf(dSquareFootage);
         	
		XSSFRow rowWindHailExclusion=sheet.getRow(a);			                                              			
		XSSFCell CollWindHailExclusion=rowWindHailExclusion.getCell(60);                       			                   			                                                       
		String WindHailExclusion=CollWindHailExclusion.getStringCellValue();
		
		XSSFRow rowGroup_ID=sheet.getRow(a);			                                              			
		XSSFCell CollGroup_ID=rowGroup_ID.getCell(61);                       			                   			                                                       
		String Group_ID=CollGroup_ID.getStringCellValue();
		
		XSSFRow rowUser_ID=sheet.getRow(a);			                                              			
		XSSFCell CollUser_ID=rowUser_ID.getCell(62);                       			                   			                                                       
		String User_ID=CollUser_ID.getStringCellValue();
		
		XSSFRow rowPassword=sheet.getRow(a);			                                              			
		XSSFCell CollPassword=rowPassword.getCell(63);                       			                   			                                                       
		String Password=CollPassword.getStringCellValue();
		
         	System.out.println(Date_OF_Birth);
         	System.out.println(New_Purchase);
	    
	     RestAssured.baseURI = "https://policy-ws.uat.thig.com";
                 Response response=
		  given()
		       .header("Content","text/xml")
	               .and()
	               .body("<soapenv:Envelope xmlns:soapenv=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:v2=\"http://www.thig.com/webservices/policy/external/v2\">\r\n" + 
	               	"   <soapenv:Header>\r\n" + 
	               	"      <v2:RequestHeader>\r\n" + 
	               	"         <v2:ApiKey>"+APIKey_UAT+"</v2:ApiKey>     \r\n" + 
	               	"      </v2:RequestHeader>\r\n" + 
	               	"   </soapenv:Header>\r\n" + 
	               	"   <soapenv:Body>\r\n" + 
	               	"      <v2:MHRateRequest>\r\n" + 
	               	"         <v2:PolicyTerm>\r\n" + 
	               	"            <v2:Coverages>\r\n" + 
	               	"               <v2:AllOtherPerilsDeductible>"+AOP+"</v2:AllOtherPerilsDeductible>\r\n" + 
	               	"               <v2:CoverageA>"+CovA+"</v2:CoverageA>\r\n" + 
	               	"               <v2:CoverageB>"+CovB+"</v2:CoverageB>\r\n" + 
	               	"               <v2:CoverageC>"+CovC+"</v2:CoverageC>\r\n" + 
	               	"               <v2:CoverageE>"+CovE+"</v2:CoverageE>\r\n" + 
	               	"               <v2:CoverageF>"+CovF+"</v2:CoverageF>\r\n" + 
	               	"               <!--Optional:-->\r\n" + 
	               	"               <v2:HurricaneDeductible>"+Hurrican_ded+"</v2:HurricaneDeductible>\r\n" + 
	               	"               <!--Optional:-->\r\n" + 
	               	"               <v2:NamedStorm>"+NamedStormDed+"</v2:NamedStorm>\r\n" + 
	               	"               <!--Optional:-->\r\n" + 
	               	"               <v2:WaterDeductible>"+Water_Ded+"</v2:WaterDeductible>\r\n" + 
	               	"               <!--Optional:-->\r\n" + 
	               	"               <v2:WindHailDeductible>"+WindHailDed+"</v2:WindHailDeductible>\r\n" + 
	               	"            </v2:Coverages>\r\n" + 
	               	"            <v2:EffectiveDate>"+EffectiveDate+"</v2:EffectiveDate>\r\n" + 
	               	"            <!--Optional:-->\r\n" + 
	               	"            <v2:Endorsements>\r\n" + 
	               	"               <!--Optional:-->\r\n" + 
	               	"               <v2:AssociationDiscount>"+AssociationType+"</v2:AssociationDiscount>\r\n" + 
	               	"               <!--Optional:-->\r\n" + 
	               	"               <v2:Paperless>"+PaperLess+"</v2:Paperless>\r\n" + 
	               	"               <!--Optional:-->\r\n" + 
	               	"               <v2:PayInFull>"+PayInFull+"</v2:PayInFull>\r\n" + 
	               	"               <!--Optional:-->\r\n" + 
	               	"               <v2:WaterDamage10AndFull>"+WaterDamage+"</v2:WaterDamage10AndFull>\r\n" + 
	               	"            </v2:Endorsements>\r\n" + 
	               	"            <v2:Form>MC</v2:Form>           \r\n" + 
	               	"            <v2:Location>\r\n" + 
	               	"               <v2:Address>\r\n" + 
	               	"                  <v2:Street>"+Address+"</v2:Street>\r\n" + 
	               	"                  \r\n" + 
	               	"                  <v2:City>"+Risk_City+"</v2:City>\r\n" + 
	               	"                  <v2:State>"+Risk_State+"</v2:State>\r\n" + 
	               	"                  <v2:Zipcode>"+Risk_ZpiCode+"</v2:Zipcode>\r\n" + 
	               	"               </v2:Address>\r\n" + 
	               	"               <v2:ConstructionYear>"+ConstructionYear+"</v2:ConstructionYear>\r\n" + 
	               	"               <v2:DwellingSettlementType>"+DwellingLossSettlement+"</v2:DwellingSettlementType>             \r\n" + 
	               	"               <!--Optional:-->\r\n" + 
	               	"               <v2:MHAdvantageChoiceHome>"+MH_AdvantageHome+"</v2:MHAdvantageChoiceHome>\r\n" + 
	               	"               <!--Optional:-->\r\n" + 
	               	"               <v2:MobileHomeType>"+MobileHomeType+"</v2:MobileHomeType>\r\n" + 
	               	"               <!--Optional:-->\r\n" + 
	               	"               <v2:MonthsUnoccupied>"+MonthsUnOccupied+"</v2:MonthsUnoccupied>\r\n" + 
	               	"               <v2:NearFireHydrant>"+Near_Hyderate+"</v2:NearFireHydrant>\r\n" + 
	               	"               <v2:Occupancy>"+Occupancy+"</v2:Occupancy>\r\n" + 
	               	"               <v2:OccupiedBy>"+OccupiedBY+"</v2:OccupiedBy>\r\n" + 
	               	"               <!--Optional:-->\r\n" + 
	               	"               <v2:ParkStatus>"+ParkStatus+"</v2:ParkStatus>\r\n" + 
	               	"               <v2:RoofYear>"+RooFYear+"</v2:RoofYear>\r\n" + 
	               	"               <!--Optional:-->\r\n" + 
	               	"               <v2:ShortTermRentalSurcharge>"+ShortTermRentalSurecharge+"</v2:ShortTermRentalSurcharge>     \r\n" + 
	               	"            </v2:Location>\r\n" + 
	               	"            <v2:PersonalPropertyReplacementCost>"+PersonalPropertyReplacementCost+"</v2:PersonalPropertyReplacementCost>\r\n" + 
	               	"            <v2:PrimaryInsured>\r\n" + 
	               	"               <!--Optional:-->\r\n" + 
	               	"               <v2:EmailAddress>"+Email+"</v2:EmailAddress>\r\n" + 
	               	"               <!--Optional:-->\r\n" + 
	               	"               <v2:InsuranceScore>"+InsurenceScore+"</v2:InsuranceScore>\r\n" + 
	               	"               <!--Optional:-->\r\n" + 
	               	"               <v2:MailingAddress>\r\n" + 
	               	"                 \r\n" + 
	               	"                  <!--Optional:-->\r\n" + 
	               	"                  <v2:Address>\r\n" + 
	               	"                     <v2:Street>"+Mailing_Street+"</v2:Street>\r\n" + 
	               	"                     <v2:City>"+Mailing_City+"</v2:City>\r\n" + 
	               	"                     <v2:State>"+Mailing_State+"</v2:State>\r\n" + 
	               	"                     <v2:Zipcode>"+MailingZipcode_ZpiCode+"</v2:Zipcode>\r\n" + 
	               	"                  </v2:Address>\r\n" + 
	               	"               </v2:MailingAddress>\r\n" + 
	               	"               <v2:FirstName>"+FirstName+"</v2:FirstName>              \r\n" + 
	               	"               <v2:LastName>"+LastName+"</v2:LastName>\r\n" + 
	               	"               <v2:DateOfBirth>"+Date_OF_Birth+"</v2:DateOfBirth>\r\n" + 
	               	"               <!--Optional:-->\r\n" + 
	               	"               <v2:PreviousAddress>                 \r\n" + 
	               	"                  <!--Optional:-->\r\n" + 
	               	"                  <v2:Address>\r\n" + 
	               	"                     <v2:Street>"+Previous_Street+"</v2:Street>\r\n" + 
	               	"                     <v2:City>"+Previous_City+"</v2:City>\r\n" + 
	               	"                     <v2:State>"+Previous_State+"</v2:State>\r\n" + 
	               	"                     <v2:Zipcode>"+PreviousZipcode_ZpiCode+"</v2:Zipcode>\r\n" + 
	               	"                  </v2:Address>\r\n" + 
	               	"               </v2:PreviousAddress>\r\n" + 
	               	"            </v2:PrimaryInsured>\r\n" + 
	               	"            <v2:Underwriting>\r\n" + 
	               	"               <!--Optional:-->\r\n" + 
	               	"               <v2:NewPurchase>"+New_Purchase+"</v2:NewPurchase>\r\n" + 
	               	"               <!--Optional:-->\r\n" + 
	               	"               <v2:OccasionalRentalSurcharge>"+OccasioanlRental_Surecharge+"</v2:OccasionalRentalSurcharge>\r\n" + 
	               	"               <!--Optional:-->\r\n" + 
	               	"               <v2:PriorPolicyExpirationDate>"+Prior_exp_Date+"</v2:PriorPolicyExpirationDate>\r\n" + 
	               	"               <v2:SquareFootage>"+SquareFootage+"</v2:SquareFootage>\r\n" + 
	               	"             </v2:Underwriting>\r\n" + 
	               	"            <v2:WindstormHailExclusion>"+WindHailExclusion+"</v2:WindstormHailExclusion>\r\n" + 
	               	"         </v2:PolicyTerm>\r\n" + 
	               	"         <v2:User>\r\n" + 
	               	"            <v2:GroupId>"+Group_ID+"</v2:GroupId>\r\n" + 
	               	"            <v2:UserId>"+User_ID+"</v2:UserId>\r\n" + 
	               	"            <v2:Password></v2:Password>\r\n" + 
	               	"         </v2:User>\r\n" + 
	               	"      </v2:MHRateRequest>\r\n" + 
	               	"   </soapenv:Body>\r\n" + 
	               	"</soapenv:Envelope>")
	        .when()
	                .post("/v2/PolicyService")
	        .then()  	               
	               // .statusCode(200) 
	                .and()
	                .log().all().extract().response();                 
                 System.out.println(response.getStatusCode());		
                 String stringResponse = response.asString();
	         XmlPath xmlpath = new XmlPath(stringResponse);	       
	         String PolicyNumber= xmlpath.getString("Envelope.Body.MHRateResponse.RateResults.RateResult.PolicyTerm.PolicyNumber");  
	         String Premium= xmlpath.getString("Envelope.Body.MHRateResponse.RateResults.RateResult.PolicyTerm.Premiums.TotalPremium");
	         String Policy_Form= xmlpath.getString("Envelope.Body.MHRateResponse.RateResults.RateResult.PolicyTerm.Form");
	         String ID_Generated= xmlpath.getString("Envelope.Body.MHRateResponse.RateResults.RateResult.PolicyTerm.Id");
	         System.out.println("Premium  :"+ "  "+Premium);
	         System.out.println("PolicyNumber :"+ "  "+PolicyNumber);
	         System.out.println("PolicyNumber :"+ "  "+Policy_Form);
	         System.out.println("PolicyNumber :"+ "  "+ID_Generated);         
	         System.out.println("next Row   :" + (a+1));
	         File file = new File("C:\\Users\\vdaru\\Desktop\\API_Automation\\MH\\API_MH.xlsx");	
			try {
				FileInputStream fis11=new FileInputStream(file);	
				wb=new XSSFWorkbook(fis11);
				sheet=wb.getSheetAt(0);	
				sheet.getRow(a).createCell(64).setCellValue(PolicyNumber);
				sheet.getRow(a).createCell(65).setCellValue(Premium);
				sheet.getRow(a).createCell(66).setCellValue(Policy_Form);
				sheet.getRow(a).createCell(67).setCellValue(ID_Generated);
				FileOutputStream fos=new FileOutputStream(file);			
			wb.write(fos);
			fos.close();
			}		
			catch(FileNotFoundException e1) {
				// TODO Auto-generated catch block		
				e1.printStackTrace();
			}	     		         
			 }	   
		}
    }
}

