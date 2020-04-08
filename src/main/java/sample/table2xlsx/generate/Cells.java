package sample.table2xlsx.generate;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.util.NumberUtils;
import org.springframework.util.StringUtils;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.temporal.Temporal;
import java.util.Map;

import static org.apache.poi.ss.usermodel.Row.MissingCellPolicy.CREATE_NULL_AS_BLANK;

public class Cells {
  
  private Cell cell;
  
  public Cells(Cell cell) {
    this.cell = cell;
  }
  
  
  public Cells value(String text) {
    cell.setCellValue(text);
    return this;
  }
  
  public Object value(final Object value, final Map<StyleType, CellStyle> styles) {
    if (StringUtils.isEmpty(value)) {
      cell.setCellValue("");
      style(styles.get(StyleType.BASE));
    }
    else if (value instanceof String) {
      cell.setCellValue(value.toString());
      style(styles.get(StyleType.BASE));
    }
    else if (value instanceof Number) {
      cell.setCellValue(NumberUtils.convertNumberToTargetClass((Number) value, Double.class));
      style(styles.get(StyleType.NUMBER));
    }
    else if (value.getClass().isPrimitive()) {
      cell.setCellValue(Double.parseDouble(value.toString()));
      style(styles.get(StyleType.NUMBER));
    }
    else if (value instanceof LocalDateTime) {
      cell.setCellValue((LocalDateTime) value);
    }
    else if (value instanceof LocalDate) {
      cell.setCellValue((LocalDate) value);
    }
    
    return this;
  }
  
  public Object value() {
    switch (cell.getCellType()) {
      case STRING:
        return cell.getStringCellValue();
      case _NONE:
        break;
      case NUMERIC:
        return DateUtil.isCellDateFormatted(cell)
            ? cell.getLocalDateTimeCellValue()
            : cell.getNumericCellValue();
      case BOOLEAN:
        return cell.getBooleanCellValue();
      case BLANK:
        return "";
      case FORMULA:
      case ERROR:
        throw new UnsupportedOperationException();
    }
    
    return null;
  }
  
  public Cells toX(final int x) {
    cell = cell.getRow().getCell(x, CREATE_NULL_AS_BLANK);
    return this;
  }
  
  public Cells next() {
    return toX(cell.getColumnIndex() + 1);
  }
  
  public Cells toY(final int y) {
    final Sheet sheet = cell.getRow().getSheet();
    Row row = sheet.getRow(y);
    if (row == null) {
      row = sheet.createRow(y);
    }
    
    cell = row.getCell(cell.getColumnIndex(), CREATE_NULL_AS_BLANK);
    return this;
  }
  
  public Cells style(final CellStyle style) {
    this.cell.setCellStyle(style);
    return this;
  }
  
  public Cells merge(int rows, int cols) {
    if (rows <= 1 && cols <= 1) {
      return this;
    }
    
    final int lastRow = cell.getRowIndex() + rows - 1;
    final int lastCol = cell.getColumnIndex() + cols - 1;
    final Sheet sheet = cell.getRow().getSheet();
    final CellStyle currentStyle = cell.getCellStyle();
    
    for (int y = cell.getRowIndex(); y < lastRow; y++) {
      for (int x = cell.getColumnIndex(); x < lastCol; x++) {
        new CellHandler(sheet).cell(y, x).style(currentStyle);
      }
    }
    
    final CellRangeAddress region = new CellRangeAddress(
        cell.getRowIndex(),
        lastRow,
        cell.getColumnIndex(),
        lastCol);
    
    sheet.addMergedRegion(region);
    
    return this;
  }
  
}
