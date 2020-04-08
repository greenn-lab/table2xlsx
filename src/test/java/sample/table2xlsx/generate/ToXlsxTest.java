package sample.table2xlsx.generate;

import com.fasterxml.jackson.databind.ObjectMapper;
import org.junit.jupiter.api.Test;
import sample.table2xlsx.ToXlsx;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

class ToXlsxTest {
  
  private String headHtml =
      "<table border=1><thead>\n" +
          "<tr><th rowspan=3>number</th><th colspan=3>name</th><th rowspan=2 colspan=3>age</th></tr>\n" +
          "<tr><th colspan=2>full name</th><th rowspan=2>middle name</th></tr>\n" +
          "<tr><th>first name</th><th>last name</th><th>day</th><th>month</th><th>year</th></tr>\n" +
          "</thead><table>";
  
  private String jsonData = "[\n" +
      "  {\n" +
      "  \"no\": 1,\n" +
      "  \"firstName\": \"Green\",\n" +
      "  \"lastName\": \"Bak\",\n" +
      "  \"middleName\": \"-\",\n" +
      "  \"day\": 31,\n" +
      "  \"month\": \"april\",\n" +
      "  \"year\": 2020\n" +
      "},{\n" +
      "  \"no\": 2,\n" +
      "  \"firstName\": \"eugene\",\n" +
      "  \"lastName\": \"cha\",\n" +
      "  \"middleName\": \"Liz\",\n" +
      "  \"day\": 31,\n" +
      "  \"month\": \"april\",\n" +
      "  \"year\": 2020\n" +
      "}\n" +
      "]";
  
  @Test
  void shouldCreateXlsxFile() throws IOException {
    
    @SuppressWarnings({"unchecked"})
    final List<Map<String, Object>> data =
        new ObjectMapper().readValue(jsonData, List.class);
    
    final List<String> orderedKeys = Arrays.asList("no", "firstName", "lastName", "middleName", "day", "month", "year");
    
    ToXlsx.builder()
        .headTableTag(headHtml)
        .out(new FileOutputStream("/Users/green/Desktop/x.xlsx"))
        .build()
        .output("hello", data, orderedKeys);
  }
}
