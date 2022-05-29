package fr.enit.industryportal.fairtoexcel.converters;

import com.google.gson.JsonElement;
import com.google.gson.JsonObject;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.awt.*;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

public abstract class ExcelConverter {

    protected List<String> headers = new ArrayList<>();

    protected JsonObject jsonSource;


    public ExcelConverter(JsonObject jsonSource) {
        this.jsonSource = jsonSource;
    }


    public void toExcel(String filePath, String sheetName) throws IOException, InvalidFormatException {
        this.toExcel(filePath, sheetName, false);
    }

    public void toExcel(String filePath, String sheetName, boolean foreNewFile) throws IOException, InvalidFormatException {
        Path path = this.getFileDir(filePath);
        if (path != null)
            Files.createDirectories(this.getFileDir(filePath));

        File file = new File(filePath);

        if (foreNewFile)
            file.delete();

        XSSFWorkbook APWorkbook = getWorkBook(file);

        XSSFSheet sheet = APWorkbook.createSheet(sheetName);
        System.out.println("Sheet '" + sheetName + "' created ");
        fillContent(APWorkbook, sheet, jsonSource);
        fillHeader(APWorkbook, sheet);

        for (int i = 0; i < getHeaders().length; i++) {
            sheet.autoSizeColumn(i);
        }


        FileOutputStream APFileOut = new FileOutputStream(filePath, true);
        APWorkbook.write(APFileOut);
        APFileOut.close();
        APWorkbook.close();
    }

    protected XSSFWorkbook getWorkBook(File file) throws IOException {
        XSSFWorkbook workbook;
        if (file.exists()) {
            workbook = (XSSFWorkbook) WorkbookFactory.create(file);
        } else {
            workbook = new XSSFWorkbook();
        }
        return workbook;
    }

    protected abstract void fillContent(XSSFWorkbook workbook, XSSFSheet sheet, JsonObject jsonSource);


    protected void fillHeader(XSSFWorkbook workbook, XSSFSheet sheet) {
        CellStyle headerCellStyle = getHeaderStyle(workbook);

        String[] headers = getHeaders();
        Row headerRow = sheet.createRow(0);

        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerCellStyle);
        }
    }


    protected Path getFileDir(String filePath) {
        int index = filePath.lastIndexOf(File.separator);
        if (index >= 0)
            return Path.of(filePath.substring(0, index));
        else
            return null;
    }

    protected double getAsDouble(String fieldName, JsonElement object) {
        return Math.round(object.getAsJsonObject().get(fieldName).getAsDouble() * 100.0) / 100.0;
    }

    protected JsonObject filterKeys(JsonElement object) {
        List<String> keysToFilter = Arrays.asList("score", "normalizedScore", "maxCredits", "maxCredits", "portalMaxCredits", "executionTime");
        JsonObject out = new JsonObject();

        for (String s : object.getAsJsonObject().keySet().stream()
                .filter(key -> !keysToFilter.contains(key))
                .collect(Collectors.toList())) {
            out.add(s, object.getAsJsonObject().get(s));
        }
        return out;
    }

    protected int getNormalizedTotalScore(JsonElement object) {
        return (int) getAsDouble("normalizedScore", object);
    }

    protected Cell fillCell(XSSFRow row, int colIndex, double value) {
        XSSFCell cell = row.createCell(colIndex);
        cell.setCellValue(value);
        return cell;
    }

    protected Cell fillCell(XSSFRow row, int colIndex, String value) {
        XSSFCell cell = row.createCell(colIndex);
        cell.setCellValue(value);
        return cell;
    }

    protected CellStyle getBasicStyle(XSSFWorkbook workbook) {
        XSSFFont font = workbook.createFont();
        font.setFontHeightInPoints((short) 18);
        CellStyle style = workbook.createCellStyle();
        style.setFont(font);

        return style;
    }

    protected CellStyle getLinkStyle(XSSFWorkbook workbook) {
        XSSFCellStyle hlinkstyle = workbook.createCellStyle();
        XSSFFont hlinkfont = workbook.createFont();
        hlinkfont.setUnderline(XSSFFont.U_SINGLE);
        hlinkfont.setFontHeightInPoints((short) 18);
        hlinkfont.setColor(IndexedColors.BLUE.index);
        hlinkstyle.setFont(hlinkfont);

        return hlinkstyle;
    }

    protected CellStyle getHeaderStyle(XSSFWorkbook workbook) {
        XSSFFont headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 20);

        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);
        return headerCellStyle;
    }

    protected void addHeader(int index, String value) {
        try {
            this.headers.get(index);
        } catch (IndexOutOfBoundsException e) {
            this.headers.add(value);
        }
    }

    protected String[] getHeaders() {
        return this.headers.toArray(new String[this.headers.size()]);
    }

    protected String getPortal() {
        String request = this.jsonSource.getAsJsonObject("status").get("request").getAsString();
        int index = request.indexOf("portal=");
        return request.substring(index + 7, request.indexOf("&", index + 7));
    }

    protected Hyperlink getLinkToPortal(XSSFWorkbook workbook, String key) {
        Hyperlink link = workbook.getCreationHelper().createHyperlink(HyperlinkType.URL);
        link.setAddress("http://" + getPortal() + ".lirmm.fr/ontologies/" + key);
        return link;
    }
}
