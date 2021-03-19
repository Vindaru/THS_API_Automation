import io.restassured.RestAssured;
import io.restassured.path.xml.XmlPath;
import io.restassured.response.Response;
import static io.restassured.RestAssured.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;

public class SOAP_DEMO_2 {

    //given, When ,There
    
    public static void main(String[] args) throws IOException {
	
	FileInputStream fis = new FileInputStream("C:\\Users\\vdaru\\Desktop\\API_MH.xlsx");	 
	    XSSFWorkbook wb=new XSSFWorkbook(fis);
            XSSFSheet sheet=wb.getSheet("Sheet1");	

            int noofRows = sheet.getLastRowNum();
		System.out.println("the total number of Rows are " + "------ " + noofRows);
		String[] Headers1 = new String[noofRows];
		int i, j;
		for (j = 0; j < noofRows; j++) {    			
			 for (int a = 1; a < noofRows; a++) {	
			     System.out.println("im in Row   :" + a);
	     XSSFRow rowPolicy_Number=sheet.getRow(a);			                                              			
	     XSSFCell cellPolicy_Number=rowPolicy_Number.getCell(0);                       			                   			                                                       
	     String FName=cellPolicy_Number.getStringCellValue();
	  
	     String APIKey_UAT= "6880c123d5dd4aaf96c9f8f91b4e0dab";
//	     String API_GREEN= "6880c123d5dd4aaf96c9f8f91b4e0dab"; 
//	     String API_BETA= "";
	     
	     String GroupID_TEST= "TS0A00";
	     String UserID_TEST= "Vdaru";
	     RestAssured.baseURI = "https://policy-ws.uat.thig.com";
                 Response response=
		  given()
		       .header("Content","text/xml")
	               .and()
	               .body("<soapenv:Envelope xmlns:soapenv=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:v2=\"http://www.thig.com/webservices/policy/external/v2\">\r\n" + 
	               	"   <soapenv:Header>\r\n" + 
	               	"      <v2:RequestHeader>\r\n" + 
	               	"         <v2:ApiKey>"+APIKey_UAT+"</v2:ApiKey>\r\n" + 
	               	"      </v2:RequestHeader>\r\n" + 
	               	"   </soapenv:Header>\r\n" + 
	               	"   <soapenv:Body>\r\n" + 
	               	"      <v2:MHRateRequest>\r\n" + 
	               	"         <v2:PolicyTerm>\r\n" + 
	               	"            <v2:Coverages>\r\n" + 
	               	"               <v2:AllOtherPerilsDeductible>1000</v2:AllOtherPerilsDeductible>\r\n" + 
	               	"               <v2:CoverageA>49063</v2:CoverageA>\r\n" + 
	               	"               <v2:CoverageB>20000</v2:CoverageB>\r\n" + 
	               	"               <v2:CoverageC>15000</v2:CoverageC>\r\n" + 
	               	"               <v2:CoverageE>100000</v2:CoverageE>\r\n" + 
	               	"               <v2:CoverageF>1000</v2:CoverageF>\r\n" + 
	               	"               <v2:WindHailDeductible>1000</v2:WindHailDeductible>\r\n" + 
	               	"            </v2:Coverages>\r\n" + 
	               	"            <v2:EffectiveDate>2020-05-25</v2:EffectiveDate>\r\n" + 
	               	"            <v2:Endorsements>\r\n" + 
	               	"               <v2:WaterDamage10AndFull>10pct</v2:WaterDamage10AndFull>\r\n" + 
	               	"            </v2:Endorsements>\r\n" + 
	               	"            <v2:Form>MC</v2:Form>\r\n" + 
	               	"            <v2:Location>\r\n" + 
	               	"               <v2:Address>\r\n" + 
	               	"                  <v2:Street>4713 Jobe Trl</v2:Street>\r\n" + 
	               	"                  <v2:City>Nolensville</v2:City>\r\n" + 
	               	"                  <v2:State>TN</v2:State>\r\n" + 
	               	"                  <v2:Zipcode>37135</v2:Zipcode>\r\n" + 
	               	"               </v2:Address>\r\n" + 
	               	"               <v2:ConstructionYear>2000</v2:ConstructionYear>\r\n" + 
	               	"               <v2:DwellingSettlementType>RC</v2:DwellingSettlementType>\r\n" + 
	               	"               <!--Optional:-->\r\n" + 
	               	"               <v2:MobileHomeType>SingleWide</v2:MobileHomeType>\r\n" + 
	               	"               <!--Optional:-->\r\n" + 
	               	"               <v2:MonthsUnoccupied>1</v2:MonthsUnoccupied>\r\n" + 
	               	"               <v2:NearFireHydrant>true</v2:NearFireHydrant>\r\n" + 
	               	"               <v2:Occupancy>Seasonal</v2:Occupancy>\r\n" + 
	               	"               <v2:OccupiedBy>Owner</v2:OccupiedBy>\r\n" + 
	               	"               <!--Optional:-->\r\n" + 
	               	"               <v2:ParkStatus>InPark26Plus</v2:ParkStatus>\r\n" + 
	               	"               <v2:RoofYear>2000</v2:RoofYear>\r\n" + 
	               	"            </v2:Location>\r\n" + 
	               	"            <v2:PersonalPropertyReplacementCost>true</v2:PersonalPropertyReplacementCost>\r\n" + 
	               	"            <v2:PrimaryInsured>\r\n" + 
	               	"               <v2:FirstName>"+FName+"</v2:FirstName>\r\n" + 
	               	"               <v2:LastName>BAILEY</v2:LastName>\r\n" + 
	               	"               <v2:DateOfBirth>1982-10-12</v2:DateOfBirth>\r\n" + 
	               	"            </v2:PrimaryInsured>\r\n" + 
	               	"            <v2:Underwriting>\r\n" + 
	               	"               <!--Optional:-->\r\n" + 
	               	"               <v2:NewPurchase>true</v2:NewPurchase>\r\n" + 
	               	"               <!--Optional:-->\r\n" + 
	               	"               <v2:PriorPolicyExpirationDate>2019-09-09</v2:PriorPolicyExpirationDate>\r\n" + 
	               	"               <v2:SquareFootage>1000</v2:SquareFootage>\r\n" + 
	               	"            </v2:Underwriting>\r\n" + 
	               	"            <v2:WindstormHailExclusion>false</v2:WindstormHailExclusion>\r\n" + 
	               	"         </v2:PolicyTerm>\r\n" + 
	               	"         <v2:User>\r\n" + 
	               	"            <v2:GroupId>"+GroupID_TEST+"</v2:GroupId>\r\n" + 
	               	"            <v2:UserId>"+UserID_TEST+"</v2:UserId>\r\n" + 
	               	"            <v2:Password></v2:Password>\r\n" + 
	               	"         </v2:User>\r\n" + 
	               	"      </v2:MHRateRequest>\r\n" + 
	               	"   </soapenv:Body>\r\n" + 
	               	"</soapenv:Envelope>")
	        .when()
	                .post("/v2/PolicyService")
	        .then()  	               
	                .statusCode(200) 
	                .and()
	                .log().all().extract().response();                 
                 System.out.println(response.getStatusCode());		
                 String stringResponse = response.asString();
	         XmlPath xmlpath = new XmlPath(stringResponse);
	         String AOP= xmlpath.getString("Envelope.Body.MHRateResponse.RateResults.RateResult.PolicyTerm.Coverages.AllOtherPerilsDeductible");
	         String CovarageA= xmlpath.getString("Envelope.Body.MHRateResponse.RateResults.RateResult.PolicyTerm.Coverages.CoverageA");      
	         String PolicyNumber= xmlpath.getString("Envelope.Body.MHRateResponse.RateResults.RateResult.PolicyTerm.PolicyNumber");  
	         String Premium= xmlpath.getString("Envelope.Body.MHRateResponse.RateResults.RateResult.PolicyTerm.Premiums.TotalPremium");      

	         System.out.println("AOP Value :"+"  "+AOP);	
	         System.out.println("CoverageA Value :"+"  "+CovarageA);
	         System.out.println("Premium  :"+ "  "+Premium);
	         System.out.println("AOP Value :"+ "  "+PolicyNumber);
	         
	         System.out.println("next Row   :" + (a+1));
	         File file = new File("C:\\Users\\vdaru\\Desktop\\API_MH.xlsx");	
			try {
				FileInputStream fis11=new FileInputStream(file);	
				wb=new XSSFWorkbook(fis11);
				sheet=wb.getSheetAt(0);	
				sheet.getRow(a).createCell(1).setCellValue(AOP);
				sheet.getRow(a).createCell(2).setCellValue(CovarageA);
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

