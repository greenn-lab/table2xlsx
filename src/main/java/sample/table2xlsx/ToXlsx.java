package sample.table2xlsx;

import lombok.Builder;
import lombok.RequiredArgsConstructor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;
import sample.table2xlsx.generate.CellHandler;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.List;
import java.util.Map;

@RequiredArgsConstructor
@Builder
public class ToXlsx {
  
  private static final Path XLSX_TEMPLATE;
  
  static {
    try {
      XLSX_TEMPLATE = Paths.get(new ClassPathResource("templates/excel-output-template.xlsx").getURI());
    }
    catch (IOException e) {
      throw new IllegalStateException(e);
    }
  }


  private final String headTableTag;
  private final String bodyTableTag;
  private final OutputStream out;
  
  
  public void output(final String name, final List<Map<String, Object>> data, final List<String> orderedKeys) throws IOException {
    final Path tmpXlsx = createTmpXlsx();
    
    try (final Workbook workbook = new XSSFWorkbook(tmpXlsx.toFile())) {
      final CellHandler handler = new CellHandler(workbook.createSheet(name));
 
      handler.processHead(headTableTag);
      handler.composeBody(data, orderedKeys);
      
      clearing(workbook);
    }
    catch (InvalidFormatException e) {
      throw new IOException(e);
    }
    finally {
      Files.delete(tmpXlsx);
    }
  }
  
  private Path createTmpXlsx() throws IOException {
    final Path tmpXlsx = Files.createTempFile("tmp_xlsx", "");
    Files.copy(XLSX_TEMPLATE, tmpXlsx, StandardCopyOption.REPLACE_EXISTING);
    return tmpXlsx;
  }
  
  private void clearing(Workbook workbook) throws IOException {
    workbook.removeSheetAt(0);
    workbook.write(out);
  }
  
  private void composeBody(CellHandler handle, Map<String, Object> data) {
  }
  
}
