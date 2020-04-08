package sample.table2xlsx.generate;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;

import java.util.EnumMap;
import java.util.List;
import java.util.Map;

import static org.apache.poi.ss.usermodel.Row.MissingCellPolicy.CREATE_NULL_AS_BLANK;
import static org.springframework.util.StringUtils.isEmpty;

public class CellHandler {
  
  private final Sheet sheet;
  private final Map<StyleType, CellStyle> styles = new EnumMap<>(StyleType.class);
  
  
  public CellHandler(final Sheet sheet) {
    this.sheet = sheet;
    extractTemplateStyles();
  }
  
  private void extractTemplateStyles() {
    final Sheet templateSheet = sheet.getWorkbook().getSheetAt(0);
    styles.put(StyleType.HEAD, templateSheet.getRow(0).getCell(0).getCellStyle());
    styles.put(StyleType.BASE, templateSheet.getRow(1).getCell(0).getCellStyle());
    styles.put(StyleType.NUMBER, templateSheet.getRow(1).getCell(1).getCellStyle());
    styles.put(StyleType.DATE, templateSheet.getRow(1).getCell(2).getCellStyle());
  }
  
  public Cells cell(final int rowIndex, final int colIndex) {
    Row row = sheet.getRow(rowIndex);
    if (row == null) {
      row = sheet.createRow(rowIndex);
    }
    
    return new Cells(row.getCell(colIndex, CREATE_NULL_AS_BLANK));
    
  }
  
  
  public void composeBody(List<Map<String, Object>> data, final List<String> orderedKeys) {
    int y = sheet.getLastRowNum() + 1;
    
    for (final Map<String, Object> row : data) {
      
      int x = 0;
      
      for (final String key : orderedKeys) {
        cell(y, x).value(row.get(key), styles);
        x++;
      }
      
      y++;
    }
  }
  
  public void processHead(String headTableTag) {
    final Document doc = Jsoup.parse(headTableTag);
    final CellStyle headStyle = styles.get(StyleType.HEAD);
    
    int y = 0;
    
    for (final Element tr : doc.select("tr")) {
      
      int x = 0;
      
      for (final Element th : tr.select("th")) {
        x = avoidMergedCell(y, x);
        
        cell(y, x)
            .value(th.text())
            .style(headStyle);
        
        hasMergeOptionThenMerge(th, y, x, headStyle);
        
        x++;
      }
      
      y++;
    }
  }
  
  private void hasMergeOptionThenMerge(Element th, int y, int x, CellStyle headStyle) {
    final int rowspan = isEmpty(th.attr("rowspan")) ? 0
        : Integer.parseInt(th.attr("rowspan")) - 1;
    
    final int colspan = isEmpty(th.attr("colspan")) ? 0
        : Integer.parseInt(th.attr("colspan")) - 1;
    
    if (rowspan > 0 || colspan > 0) {
      sheet.addMergedRegion(
          new CellRangeAddress(y, y + rowspan, x, x + colspan));
      
      for (int i = y; i <= y + rowspan; i++) {
        for (int j = x; j <= x + colspan; j++) {
          cell(i, j).style(headStyle);
        }
      }
    }
  }
  
  private int avoidMergedCell(int y, int x) {
    for (final CellRangeAddress address : sheet.getMergedRegions()) {
      while (address.containsRow(y) && address.containsColumn(x)) {
        x++;
      }
    }
    
    return x;
  }
  
}
