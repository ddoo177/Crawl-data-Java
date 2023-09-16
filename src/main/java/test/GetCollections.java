package test;

import com.google.gson.Gson;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GetCollections {

    public static void main(String[] args) throws IOException {

        ArrayList<Collection> listCollection = new ArrayList<>();
        Document doc = Jsoup.connect("https://yame.vn").get();

        ArrayList<String> titles = new ArrayList<>();
        ArrayList<String> collectionURLs = new ArrayList<>();

        Elements collections = doc.select("ul.dropdown");
        Elements aElements = collections.select("a");

        //Lọc qua từng thẻ a lưu href vào aElements
        for (Element aElement : aElements) {
            String groupURL = "https://yame.vn";
            groupURL = groupURL.concat(aElement.attr("href"));
            collectionURLs.add(groupURL);
            titles.add(aElement.text());
        }

        for (String collectionUrl : collectionURLs) {
            Document collectionDoc = Jsoup.connect(collectionUrl).get();
            Elements products = collectionDoc.getElementsByClass("pitem");
        
            ArrayList<String> codes = new ArrayList<>();
            Collection item = new Collection();

            for (Element productDetail : products) {
                String URL = "https://yame.vn";
                URL = URL.concat(productDetail.getElementsByTag("a").attr("href"));

                Document productDoc = Jsoup.connect(URL).get();

                //Group
                item.setTitle(collectionDoc.getElementsByClass("text-black").get(0).text().substring(11));

                Elements product = productDoc.getElementsByClass("product-info");
                for (Element e : product) {
                    String code = productDoc.getElementsByTag("p").get(0).text();
                    String[] splitCode = code.split("#");
                    codes.add(splitCode[1]);
                }
            }
            String[] arrayCode = codes.toArray(new String[0]);
            item.setProducts(arrayCode);
            listCollection.add(item);
        }
//            showJson(listCollection);
        saveExcel(listCollection);
    }

    public static void showJson(ArrayList<Collection> array) {
        Gson gson = new Gson();
        String jsonData = gson.toJson(array);;
        System.out.println(jsonData);
    }

    public static void saveExcel(ArrayList<Collection> array) {

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Collection");
        int rowNum = 0;
        Row firstRow = sheet.createRow(rowNum++);

        Cell firstCell = firstRow.createCell(0);
        firstCell.setCellValue("TITLE");
        Cell secondCell = firstRow.createCell(1);
        secondCell.setCellValue("PRODUCTS");

        for (Collection p : array) {
            Row row = sheet.createRow(rowNum++);

            Cell cell1 = row.createCell(0);
            cell1.setCellValue(p.getTitle());

            Cell cell2 = row.createCell(1);
            StringBuilder productListCell = new StringBuilder();
            for (String str : p.getProducts()) {
                productListCell.append(str).append(" ; ");
            }
            productListCell.delete(productListCell.length() - 2, productListCell.length());
            cell2.setCellValue(productListCell.toString());
        }
        try {
            FileOutputStream outputStream = new FileOutputStream("Products.xlsx");
            workbook.write(outputStream);
            workbook.close();
        } catch (FileNotFoundException e) {
            System.out.println(e);
        } catch (IOException e) {
            System.out.println(e);
        }
        System.out.println("Done");
    }
}
