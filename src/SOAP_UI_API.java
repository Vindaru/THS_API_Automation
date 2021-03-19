import io.restassured.RestAssured;
import static io.restassured.RestAssured.*;
import static org.hamcrest.Matchers.*;


public class SOAP_UI_API {

    //given, When ,There
    
    public static void main(String[] args) {
	
	RestAssured.baseURI="https://policy-ws.green.thig.com";
	given().log().all().queryParam("key", "MH")
	.body("<soapenv:Envelope xmlns:soapenv=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:v2=\"http://www.thig.com/webservices/policy/external/v2\">\r\n" + 
		"   <soapenv:Header>\r\n" + 
		"      <v2:RequestHeader>\r\n" + 
		"         <v2:ApiKey>6880c123d5dd4aaf96c9f8f91b4e0dab</v2:ApiKey>\r\n" + 
		"      </v2:RequestHeader>\r\n" + 
		"   </soapenv:Header>\r\n" + 
		"   <soapenv:Body>\r\n" + 
		"      <v2:MHRateRequest>\r\n" + 
		"         <v2:PolicyTerm>\r\n" + 
		"            <v2:Coverages>\r\n" + 
		"               <v2:AllOtherPerilsDeductible>5000</v2:AllOtherPerilsDeductible>\r\n" + 
		"               <v2:CoverageA>200000</v2:CoverageA>\r\n" + 
		"               <v2:CoverageB>60000</v2:CoverageB>\r\n" + 
		"               <v2:CoverageC>150000</v2:CoverageC>\r\n" + 
		"               <v2:CoverageE>100000</v2:CoverageE>\r\n" + 
		"               <v2:CoverageF>2000</v2:CoverageF>\r\n" + 
		"               <!--Optional:-->\r\n" + 
		"               <v2:NamedStorm>2pct</v2:NamedStorm>\r\n" + 
		"               <!--Optional:-->\r\n" + 
		"               <v2:WindHailDeductible>1000</v2:WindHailDeductible>\r\n" + 
		"            </v2:Coverages>\r\n" + 
		"            <v2:EffectiveDate>2020-05-12</v2:EffectiveDate>\r\n" + 
		"            <!--Optional:-->\r\n" + 
		"            <v2:Endorsements>\r\n" + 
		"               <!--Optional:-->\r\n" + 
		"               <v2:Paperless>false</v2:Paperless>\r\n" + 
		"               <!--Optional:-->\r\n" + 
		"               <v2:PayInFull>true</v2:PayInFull>\r\n" + 
		"            </v2:Endorsements>\r\n" + 
		"            <v2:Form>MC</v2:Form>\r\n" + 
		"            <v2:Location>\r\n" + 
		"               <v2:Address>\r\n" + 
		"                  <v2:Street>4449 W Topeka Dr</v2:Street>\r\n" + 
		"                  <v2:City>Glendale</v2:City>\r\n" + 
		"                  <v2:State>AZ</v2:State>\r\n" + 
		"                  <v2:Zipcode>85308</v2:Zipcode>\r\n" + 
		"               </v2:Address>\r\n" + 
		"               <v2:ConstructionYear>2000</v2:ConstructionYear>\r\n" + 
		"               <v2:DwellingSettlementType>RC</v2:DwellingSettlementType>\r\n" + 
		"               <!--Optional:-->\r\n" + 
		"               <v2:MobileHomeType>SingleWide</v2:MobileHomeType>\r\n" + 
		"               <!--Optional:-->\r\n" + 
		"               <v2:MonthsUnoccupied>1</v2:MonthsUnoccupied>\r\n" + 
		"               <v2:NearFireHydrant>false</v2:NearFireHydrant>\r\n" + 
		"               <v2:Occupancy>Primary</v2:Occupancy>\r\n" + 
		"               <v2:OccupiedBy>Owner</v2:OccupiedBy>\r\n" + 
		"               <!--Optional:-->\r\n" + 
		"               <v2:ParkStatus>InPark26Plus</v2:ParkStatus>\r\n" + 
		"               <v2:RoofYear>2000</v2:RoofYear>\r\n" + 
		"            </v2:Location>\r\n" + 
		"            <v2:PersonalPropertyReplacementCost>true</v2:PersonalPropertyReplacementCost>\r\n" + 
		"            <v2:PrimaryInsured>\r\n" + 
		"               <!--Optional:-->\r\n" + 
		"               <v2:FirstName>Test</v2:FirstName>\r\n" + 
		"               <v2:LastName>Testing</v2:LastName>\r\n" + 
		"               <v2:DateOfBirth>1992-02-02</v2:DateOfBirth>\r\n" + 
		"            </v2:PrimaryInsured>\r\n" + 
		"            <v2:Underwriting>\r\n" + 
		"               <!--Optional:-->\r\n" + 
		"               <v2:NewPurchase>true</v2:NewPurchase>\r\n" + 
		"               <v2:SquareFootage>1000</v2:SquareFootage>\r\n" + 
		"            </v2:Underwriting>\r\n" + 
		"            <v2:WindstormHailExclusion>false</v2:WindstormHailExclusion>\r\n" + 
		"         </v2:PolicyTerm>\r\n" + 
		"         <v2:User>\r\n" + 
		"            <v2:GroupId>TS0A00</v2:GroupId>\r\n" + 
		"            <v2:UserId>VDARU</v2:UserId>\r\n" + 
		"            <v2:Password>Dvk@1234</v2:Password>\r\n" + 
		"         </v2:User>\r\n" + 
		"      </v2:MHRateRequest>\r\n" + 
		"   </soapenv:Body>\r\n" + 
		"</soapenv:Envelope>").when().post("v2/PolicyService").then().log().all().assertThat().statusCode(200)
                      .header("Server", equalTo("Apache")).body("<wpe:AllOtherPerilsDeductible>5000</wpe:AllOtherPerilsDeductible>", equalTo("5000"));
    
    
    }
    
}
