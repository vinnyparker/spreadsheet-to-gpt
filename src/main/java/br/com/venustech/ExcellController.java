package br.com.venustech;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.*;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.client.RestTemplate;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.*;

@RestController
@RequestMapping("/api/excel")
public class ExcellController {

    private static final String CHATGPT_API_URL = "https://api.openai.com/v1/chat/completions";
    private static final String CHATGPT_API_KEY = "SUA_CHAVE_AQUI"; // manda a chave api aqui

    @PostMapping("/upload")
    public ResponseEntity<String> uploadExcel(@RequestParam("file") MultipartFile file) {
        List<Map<String, String>> dataList = extractExcelData(file);
        if (dataList == null) {
            return ResponseEntity.badRequest().body("Erro ao processar o arquivo Excel.");
        }

        // Constrói a pergunta para o ChatGPT
        String prompt = "Analise os seguintes dados e forneça insights:\n" + dataList.toString();

        // Envia a requisição para a API do ChatGPT
        String response = sendToChatGPT(prompt);

        return ResponseEntity.ok(response);
    }

    private List<Map<String, String>> extractExcelData(MultipartFile file) {
        List<Map<String, String>> dataList = new ArrayList<>();
        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {
            Sheet sheet = workbook.getSheetAt(0); // Lê a primeira aba da planilha

            Row headerRow = sheet.getRow(0);
            List<String> headers = new ArrayList<>();
            for (Cell cell : headerRow) {
                headers.add(cell.getStringCellValue());
            }

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Map<String, String> rowData = new HashMap<>();
                for (int j = 0; j < headers.size(); j++) {
                    Cell cell = row.getCell(j);
                    rowData.put(headers.get(j), getCellValueAsString(cell));
                }
                dataList.add(rowData);
            }
        } catch (IOException e) {
            return null;
        }
        return dataList;
    }

    private String getCellValueAsString(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue();
            case NUMERIC: return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN: return String.valueOf(cell.getBooleanCellValue());
            case FORMULA: return cell.getCellFormula();
            default: return "";
        }
    }

    private String sendToChatGPT(String prompt) {
        RestTemplate restTemplate = new RestTemplate();

        // Corpo da requisição para a API do ChatGPT
        String requestBody = "{"
                + "\"model\": \"gpt-4\","
                + "\"messages\": [{\"role\": \"system\", \"content\": \"Você é um assistente útil.\"},"
                + "{\"role\": \"user\", \"content\": \"" + prompt.replace("\"", "\\\"") + "\"}],"
                + "\"temperature\": 0.7"
                + "}";

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_JSON);
        headers.set("Authorization", "Bearer " + CHATGPT_API_KEY);

        HttpEntity<String> entity = new HttpEntity<>(requestBody, headers);
        ResponseEntity<String> response = restTemplate.exchange(CHATGPT_API_URL, HttpMethod.POST, entity, String.class);

        return response.getBody();
    }
}
