import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLEncoder;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import org.json.JSONArray;
import org.json.JSONObject;

public class NewExcel
{

    private String inputFile;
    private String outputFile;

    public void setInputFile(String inputFile)
    {
        this.inputFile = inputFile;
    }

    public void setOutputFile(String outputFile)
    {
        this.outputFile = outputFile;
    }

    public void read() throws IOException
    {
        File inputWorkbook = new File(inputFile);
        File outputWorkbook = new File(outputFile);
        Workbook w;
        WritableWorkbook w1 = null;
        try
        {
            w = Workbook.getWorkbook(inputWorkbook);
            w1 = Workbook.createWorkbook(outputWorkbook,w);

            // Get the first sheet
            Sheet sheet = w.getSheet(0);
            WritableSheet wsheet = w1.getSheet(0);
            // Loop over first 10 column and lines

            for(int i = 0; i < sheet.getRows(); i++)
            {
                String URL = "https://maps.googleapis.com/maps/api/geocode/json?address=";
                for (int j = 0; j < 2; j++)
                {
                    Cell cell = sheet.getCell(j, i);
                    String address  = cell.getContents();
                    address = URLEncoder.encode(address, "UTF-8");
                    URL = URL.concat(address);
                    if(j != (sheet.getColumns() - 1)){
                        URL = URL.concat(",");
                    }
                }
                URL = URL.concat("&key=AIzaSyCTfvvTi0ubKX1eew1VdLhtQC56ia66suo");
                String postal_code = "abc";
                System.out.println(URL);
                System.out.println(getHTML(URL));
                JSONObject jsonObj = new JSONObject(getHTML(URL));
                JSONArray jsonArray = jsonObj.getJSONArray("results");
                jsonObj = jsonArray.getJSONObject(0);
                jsonArray = jsonObj.getJSONArray("address_components");
                for(int k=0 ; k< jsonArray.length(); k++){
                    JSONArray jsonArray1 = jsonArray.getJSONObject(k).getJSONArray("types");
                    if(jsonArray1.getString(0).equals("postal_code")){
                        postal_code = jsonArray.getJSONObject(k).getString("long_name");
                        break;
                    }
                }
                JSONObject geometry = jsonObj.getJSONObject("geometry");
                JSONObject location = geometry.getJSONObject("location");
                Double latitude = location.getDouble("lat");
                Double longitude = location.getDouble("lng");
                Label label= new Label(sheet.getColumns(), i, postal_code);
                wsheet.addCell(label);
//                Label label1= new Label(sheet.getColumns(), i, String.valueOf(latitude));
//                wsheet.addCell(label1);
//                Label label2= new Label((sheet.getColumns() + 1), i, String.valueOf(longitude));
//                wsheet.addCell(label2);
//                System.out.println(latitude);
//                System.out.println(longitude);
                System.out.println(postal_code);
            }
            w1.write();

            /*for (int j = 0; j < sheet.getColumns(); j++)
            {
                for (int i = 0; i < sheet.getRows(); i++)
                {
                    Cell cell = sheet.getCell(j, i);
                    System.out.println(cell.getContents());
                }
            }*/
        }
        catch (BiffException e)
        {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                w1.close();
            } catch (WriteException e) {
                e.printStackTrace();
            }
        }
    }

    public static void main(String[] args) throws IOException
    {
        NewExcel test = new NewExcel();
        test.setInputFile("./data/Parking Lot - SF with Price.xls");
        test.setOutputFile("/home/affine/Output.xls");
        test.read();
    }

    public static String getHTML(String urlToRead) throws Exception {
        StringBuilder result = new StringBuilder();
        URL url = new URL(urlToRead);
        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
        conn.setRequestMethod("GET");
        BufferedReader rd = new BufferedReader(new InputStreamReader(conn.getInputStream()));
        String line;
        while ((line = rd.readLine()) != null) {
            result.append(line);
        }
        rd.close();
        return result.toString();
    }

}