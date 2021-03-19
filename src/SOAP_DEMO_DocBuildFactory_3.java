import java.io.StringReader;
 
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathFactory;
 
import org.xml.sax.InputSource;
 
public class SOAP_DEMO_DocBuildFactory_3 
{
    public static void main(String[] args) throws Exception 
    {
 
        String xml = "<employees>"
                    + "<employee id=\"1\">"
                        + "<firstName>Lokesh</firstName>"
                        + "<lastName>Gupta</lastName>"
                        + "<department><id>101</id><name>IT</name></department>"
                    + "</employee>"
                   + "</employees>";
         
        InputSource inputXML = new InputSource( new StringReader( xml ) );
         
        XPath xPath = XPathFactory.newInstance().newXPath();
         
        String result = xPath.evaluate("/employees/employee/department/id", inputXML);
 
        System.out.println(result);
    }
}    