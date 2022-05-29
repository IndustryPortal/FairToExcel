package fr.enit.industryportal.fairtoexcel.converters;

import com.google.gson.JsonElement;
import com.google.gson.JsonObject;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.util.Map;

public class QuestionsExcelConverter extends ExcelConverter{

    public QuestionsExcelConverter(JsonObject jsonSource) {
        super(jsonSource);
    }

    @Override
    protected void fillContent(XSSFWorkbook workbook, XSSFSheet sheet, JsonObject jsonSource) {

        JsonObject ontologies = jsonSource.getAsJsonObject("ontologies");
        CellStyle basicCellStyle = getBasicStyle(workbook);
        CellStyle linkCellStyle = getLinkStyle(workbook);
        XSSFRow row;
        Cell cell;
        int rowNum = 1, colNum;

        for (Map.Entry<String, JsonElement> onto : ontologies.entrySet()) {
            System.out.println("> In Sheet <"+ sheet.getSheetName() +"> Writing  row " + (rowNum) + " of " + ontologies.size() + " the fair scores of : " + onto.getKey() );

            row = sheet.createRow(rowNum);

            addHeader(0 , "Ontologies");
            cell = fillCell(row , 0 ,onto.getKey());
            cell.setCellStyle(linkCellStyle);
            cell.setHyperlink(getLinkToPortal(workbook,onto.getKey()));


            fillCell(row , 1 , getNormalizedTotalScore(onto.getValue())).setCellStyle(basicCellStyle);
            addHeader(1 , "Fair score");

            colNum = 2;
            JsonObject principals = filterKeys(onto.getValue());
            for (Map.Entry<String, JsonElement> p : principals.entrySet()) {

                fillCell(row ,colNum ,getNormalizedTotalScore(p.getValue())).setCellStyle(basicCellStyle);
                addHeader(colNum , p.getKey());
                colNum++;

                JsonObject criteria = filterKeys(p.getValue());
                for (Map.Entry<String, JsonElement> c : criteria.entrySet()) {
                    fillCell(row , colNum , getNormalizedTotalScore(c.getValue())).setCellStyle(basicCellStyle);
                    addHeader(colNum , c.getKey());
                    JsonObject questions = c.getValue().getAsJsonObject().getAsJsonObject("results");
                    colNum++;
                    for (Map.Entry<String, JsonElement> question : questions.entrySet()) {
                        fillCell(row, colNum , getAsDouble("score" , question.getValue())).setCellStyle(basicCellStyle);
                        addHeader(colNum , question.getKey());

                        colNum++;
                    }
                }
            }
            rowNum++;
        }
    }


}
