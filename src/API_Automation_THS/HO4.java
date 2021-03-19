package API_Automation_THS;

import io.restassured.RestAssured;
import io.restassured.path.xml.XmlPath;
import io.restassured.response.Response;
import static io.restassured.RestAssured.*;
import java.io.FileOutputStream;
import Applicability_Package.Constant;
import Applicability_Package.MH_ExcelUtils;


public class HO4  extends MH_ExcelUtils{

    //given, When ,There
    
    public static void main(String[] args) throws Exception {
           
     MH_ExcelUtils.setExcelFile(Constant.HO4_Path_TestData + Constant.HO4_File_TestData, "sheet1");		
	int i, j;
	for (j = 0; j < MH_ExcelUtils.ExcelWSheet.getPhysicalNumberOfRows(); j++) {    			
	    for (int a = 3; a <  MH_ExcelUtils.ExcelWSheet.getPhysicalNumberOfRows(); a++) {	
			
		// read values form EXCEL;    	     
	          String APIKey_UAT =                                                        MH_ExcelUtils.getCellData(a, 1);
		  int AOP =                                       (int)                      MH_ExcelUtils.getNumericCellValue(a, 2);  		     
		  int CovA =                                      (int)                      MH_ExcelUtils.getNumericCellValue(a, 3);       
		  int CovB =                                      (int)                      MH_ExcelUtils.getNumericCellValue(a, 4);       
		  int CovC =                                      (int)                      MH_ExcelUtils.getNumericCellValue(a, 5);  	     
		  int CovE =                                      (int)                      MH_ExcelUtils.getNumericCellValue(a, 6);  
		  int CovF =                                      (int)                      MH_ExcelUtils.getNumericCellValue(a, 7);  
		  String NamedStormDed =                                                     MH_ExcelUtils.getCellData(a, 9);
		  int Hurrican_ded =                              (int)                      MH_ExcelUtils.getNumericCellValue(a, 8);  
		  int Water_Ded =                                 (int)                      MH_ExcelUtils.getNumericCellValue(a, 10);  
		  int WindHailDed =                               (int)                      MH_ExcelUtils.getNumericCellValue(a, 11);  
		  String EffectiveDate =                                                     MH_ExcelUtils.getCellData(a, 12);
		  String AssociationType =                                                   MH_ExcelUtils.getCellData(a, 13);
		  String PaperLess =                                                         MH_ExcelUtils.getCellData(a, 14);
		  String PayInFull =                                                         MH_ExcelUtils.getCellData(a, 15);
		  String WaterDamage =                                                       MH_ExcelUtils.getCellData(a, 16);
		  String Address =                                                           MH_ExcelUtils.getCellData(a, 18);
		  String Risk_City =                                                         MH_ExcelUtils.getCellData(a, 19);
		  String Risk_State =                                                        MH_ExcelUtils.getCellData(a, 20);
		  int Risk_ZpiCode =                              (int)                      MH_ExcelUtils.getNumericCellValue(a, 21); 
		  int ConstructionYear =                          (int)                      MH_ExcelUtils.getNumericCellValue(a, 22); 
		  String DwellingLossSettlement =                                            MH_ExcelUtils.getCellData(a, 23);
		  String MH_AdvantageHome =                                                  MH_ExcelUtils.getCellData(a, 24);
		  String MobileHomeType =                                                    MH_ExcelUtils.getCellData(a, 25);
		  int MonthsUnOccupied =                          (int)                      MH_ExcelUtils.getNumericCellValue(a, 26); 
		  String Near_Hyderate =                                                     MH_ExcelUtils.getCellData(a, 27);
		  String Occupancy =                                                         MH_ExcelUtils.getCellData(a, 28);
		  String OccupiedBY =                                                        MH_ExcelUtils.getCellData(a, 29);
		  String ParkStatus =                                                        MH_ExcelUtils.getCellData(a, 30);
		  int RooFYear =                                  (int)                      MH_ExcelUtils.getNumericCellValue(a, 31); 
		  String ShortTermRentalSurecharge =                                         MH_ExcelUtils.getCellData(a, 32);
		  int ProtectionClass =                           (int)                      MH_ExcelUtils.getNumericCellValue(a, 33); 
		  String PersonalPropertyReplacementCost =                                   MH_ExcelUtils.getCellData(a, 34);
		  String Email =                                                             MH_ExcelUtils.getCellData(a, 35);
		  String InsurenceScore =                                                    MH_ExcelUtils.getCellData(a, 36);
		  String Mailing_County =                                                    MH_ExcelUtils.getCellData(a, 37);
		  String Mailing_Address =                                                   MH_ExcelUtils.getCellData(a, 39);
		  String Mailing_City =                                                      MH_ExcelUtils.getCellData(a, 40);
		  String Mailing_State =                                                     MH_ExcelUtils.getCellData(a, 41);
		  int MailingZipcode_ZpiCode =                    (int)                      MH_ExcelUtils.getNumericCellValue(a, 42);
		  String FirstName =                                                         MH_ExcelUtils.getCellData(a, 43);
		  String LastName =                                                          MH_ExcelUtils.getCellData(a, 44);
		  String Date_OF_Birth =                                                     MH_ExcelUtils.getCellData(a, 45);
		  String Previous_County =                                                   MH_ExcelUtils.getCellData(a, 46);
		  String Previous_Address =                                                  MH_ExcelUtils.getCellData(a, 48);
		  String Previous_City =                                                     MH_ExcelUtils.getCellData(a, 49);
		  String Previous_State =                                                    MH_ExcelUtils.getCellData(a, 50);
		  int PreviousZipcode_ZpiCode =                   (int)                      MH_ExcelUtils.getNumericCellValue(a, 51);
		  String New_Purchase =                                                      MH_ExcelUtils.getCellData(a, 52);
		  String OccasioanlRental_Surecharge =                                       MH_ExcelUtils.getCellData(a, 53);
		  String Prior_exp_Date =                                                    MH_ExcelUtils.getCellData(a, 54);
		  int SquareFootage =                             (int)                      MH_ExcelUtils.getNumericCellValue(a, 55);
		  String WindHailExclusion =                                                 MH_ExcelUtils.getCellData(a, 60);
		  String Group_ID =                                                          MH_ExcelUtils.getCellData(a, 61);
		  String User_ID =                                                           MH_ExcelUtils.getCellData(a, 62);
		  String Password =                                                          MH_ExcelUtils.getCellData(a, 63);
		 
   
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
	               	"                     <v2:Street>"+Mailing_Address+"</v2:Street>\r\n" + 
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
	               	"                     <v2:Street>"+Previous_Address+"</v2:Street>\r\n" + 
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
	         System.out.println("next Row   :" + (a + 1));
	       
  // Write in Excel 
	         MH_ExcelUtils.setExcelFile(Constant.HO4_Path_TestData + Constant.HO4_File_TestData, "sheet1"); 
                     MH_ExcelUtils.ExcelWSheet.getRow(a).createCell(1).setCellValue(PolicyNumber);
		     MH_ExcelUtils.ExcelWSheet.getRow(a).createCell(2).setCellValue(Premium);
	             MH_ExcelUtils.ExcelWSheet.getRow(a).createCell(3).setCellValue(Policy_Form);
	             MH_ExcelUtils.ExcelWSheet.getRow(a).createCell(4).setCellValue(ID_Generated);
	             FileOutputStream fos=new FileOutputStream(Constant.HO4_Path_TestData + Constant.HO4_File_TestData);			
	          ExcelWBook.write(fos);
                  fos.close();
			     		         
			 }	   
		}
    }
}


