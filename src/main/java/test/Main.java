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
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {

    public static void main(String[] args) throws IOException {

        ArrayList<Product> listProduct = new ArrayList<>();
        Document doc = Jsoup.connect("https://yame.vn").get();

        ArrayList<String> groupURLs = new ArrayList<>();
        ArrayList<String> categories = new ArrayList<>();
        Element group = doc.select("div.dropdown").first();

        Elements aElements = group.select("a");
        //Lọc qua từng thẻ a lưu href vào aElements
        for (Element aElement : aElements) {
            if (aElement.text().equals("Áo Thun") || aElement.text().equals("Áo Sơ Mi")
                    || aElement.text().equals("Áo Khoác") || aElement.text().equals("Phụ kiện")
                    || aElement.text().equals("Sale Giày Dép SX 2022")) {
                continue;
            } else {
                String groupURL = "https://yame.vn";
                groupURL = groupURL.concat(aElement.attr("href"));
                groupURLs.add(groupURL);
                categories.add(aElement.text());
            }
        }
        System.out.println(groupURLs.size());

        for (String collectionURL : groupURLs) {
            Document collectionDoc = Jsoup.connect(collectionURL).get();
            Elements products = collectionDoc.getElementsByClass("pitem");

            for (String category : categories) {
                for (Element productDetail : products) {
                    Product item = new Product();
                    item.setCategory(category);
                    String productUrl = "https://yame.vn";
                    productUrl = productUrl.concat(productDetail.getElementsByTag("a").attr("href"));
                    item.setUrl(productUrl);
                    Document productDoc = Jsoup.connect(productUrl).get();

                    Elements product = productDoc.getElementsByClass("product-info");
                    for (Element e : product) {
                        //Name
                        item.setTitle(e.getElementsByClass("productName").text());

                        //Code
                        String code = e.getElementsByTag("p").get(0).text();
                        String[] splitCode = code.split("#");
                        item.setCode(splitCode[1]);

                        //Price
                        String priceString = e.getElementsByClass("price").get(0).text();
                        priceString = priceString.replaceAll("[\\,\\đ\\.\\^:, ]", "");
                        long priceNumber = Long.parseLong(priceString);
                        item.setPrice(priceNumber);

                        // Image URL
                        item.setMainImageUrl(e.getElementsByTag("img").attr("src"));

                        //Describe
                        try {
                            String productDesc = e.getElementById("moTaSanPham").html();
                            String[] lines = productDesc.split("<br>");
                            String describe = "";
                            for (String line : lines) {
                                if (line.contains("Mô tả sản phẩm")) {
                                    continue;
                                } else {
                                    describe = describe.concat(line).trim();
                                }
                            }
                            item.setDescribe(describe);
                        } catch (Exception ex) {
                            System.out.println(ex);
                        }

                        //Attributes
                        Element attriTable = productDoc.select("table").first();
                        Elements rows = attriTable.select("tr");
                        String[] arrayAttributes;
                        ArrayList<String> listAttributes = new ArrayList<>();
                        for (Element row : rows) {
                            Elements columns = row.select("td");
                            if (!columns.isEmpty()) {
                                Element firstColumn = columns.get(0);
                                String output = firstColumn.text();
                                listAttributes.add(output);
                            }
                        }
                        arrayAttributes = listAttributes.toArray(new String[0]);
                        item.setAttributes(arrayAttributes);

                        //Image List Url;
                        Elements imageList = productDoc.getElementsByClass("detailImageItem");
                        String[] arrayImages;
                        ArrayList<String> listImages = new ArrayList<>();
                        for (Element image : imageList) {
                            listImages.add(image.attr("src"));
                        }
                        arrayImages = listImages.toArray(new String[0]);
                        item.setImageListUrl(arrayImages);
                        listProduct.add(item);
                    }
                }
            }
        }
//        showJson(listProduct);
        saveExcel(listProduct);
    }

    public static void showJson(ArrayList<Product> array) {
        Gson gson = new Gson();
        String jsonData = gson.toJson(array);;
        System.out.println(jsonData);
    }

    public static void saveExcel(ArrayList<Product> array) {

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Product");
        int rowNum = 0;
        Row firstRow = sheet.createRow(rowNum++);

        Cell firstCell = firstRow.createCell(0);
        firstCell.setCellValue("CODE");
        Cell secondCell = firstRow.createCell(1);
        secondCell.setCellValue("TITLE");
        Cell thirdCell = firstRow.createCell(2);
        thirdCell.setCellValue("PRICE");
        Cell forthCell = firstRow.createCell(3);
        forthCell.setCellValue("GROUP");
        Cell fifthCell = firstRow.createCell(4);
        fifthCell.setCellValue("CATEGORY");
        Cell sixthCell = firstRow.createCell(5);
        sixthCell.setCellValue("ITEM_URL");
        Cell seventhCell = firstRow.createCell(6);
        seventhCell.setCellValue("MAIN_IMAGE_URL");
        Cell eighthCell = firstRow.createCell(7);
        eighthCell.setCellValue("DESCRIBE");
        Cell ninethCell = firstRow.createCell(8);
        ninethCell.setCellValue("ATRIBUTES");
        Cell tenthCell = firstRow.createCell(9);
        tenthCell.setCellValue("IMAGE_LIST_URL");

        for (Product p : array) {
            Row row = sheet.createRow(rowNum++);

            Cell cell1 = row.createCell(0);
            cell1.setCellValue(p.getCode());

            Cell cell2 = row.createCell(1);
            cell2.setCellValue(p.getTitle());

            Cell cell3 = row.createCell(2);
            cell3.setCellValue(p.getPrice());

            Cell cell4 = row.createCell(3);
            cell4.setCellValue(p.getGroup());

            Cell cell5 = row.createCell(4);
            cell5.setCellValue(p.getCategory());

            Cell cell6 = row.createCell(5);
            cell6.setCellValue(p.getUrl());

            Cell cell7 = row.createCell(6);
            cell7.setCellValue(p.getMainImageUrl());

            Cell cell8 = row.createCell(7);
            cell8.setCellValue(p.getDescribe());

            Cell cell9 = row.createCell(8);
            StringBuilder attributeCell = new StringBuilder();
            for (String str : p.getAttributes()) {
                attributeCell.append(str).append(" ; ");
            }
            attributeCell.delete(attributeCell.length() - 2, attributeCell.length());
            cell9.setCellValue(attributeCell.toString());

            Cell cell10 = row.createCell(9);
            StringBuilder imageListUrlCell = new StringBuilder();
            for (String str : p.getImageListUrl()) {
                imageListUrlCell.append(str).append(" ; ");
            }
            imageListUrlCell.delete(imageListUrlCell.length() - 2, imageListUrlCell.length());
            cell10.setCellValue(imageListUrlCell.toString());
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
