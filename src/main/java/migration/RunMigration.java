package migration;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.util.*;

public class RunMigration {

    private static final String DIR_PATH = "";

    public static void main(String[] args) {

        System.out.println("LOG -> Lendos todos os arquivos .json..");
        List<File> files = new ArrayList<>();
        listf(DIR_PATH, files);
        System.out.println("LOG -> Lista de arquivos recuperadas com sucesso!");

        for (File file : files) {
            try {
                System.out.println("---------------------------------------------------------------------------------");
                runConverter(file);
            }catch (Exception e){
                System.err.println("ERR -> Ocorreu um erro inesperado: " + e.getMessage());
            }
        }

        System.out.println("LOG -> Convesão concluidao com sucesso!");
    }

    private static void runConverter(File file) throws Exception{

        System.out.println("LOG -> Iniciando conversao de: " + file.getAbsolutePath());
        List<Map<String, Object>> data;

        try {
            System.out.println("LOG -> Lendo conteudo do arquivo..");
            String json = Files.readString(file.toPath());
            ObjectMapper mapper = new ObjectMapper();
            data = mapper.readValue(json, new TypeReference<>() {});
            System.out.println("LOG -> Arquivo lido com sucesso!");
        } catch (IOException e) {
            e.printStackTrace();
            return;
        }

        System.out.println("LOG -> Gerando arquivo xlsx..");
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet(UUID.randomUUID().toString());

        Map<String, Integer> mapHeaders = new LinkedHashMap<>();
        convertHeader(data, mapHeaders);
        
        int cellNum = 0;
        Row headerRow = sheet.createRow(0);
        for (String key : mapHeaders.keySet()) {
            Cell cell = headerRow.createCell(cellNum++);
            cell.setCellValue(key);
        }

        int rowNum = 1;
        for (Map<String, Object> obj : data) {
            Row row = sheet.createRow(rowNum++);
            for (String key : obj.keySet()) {
                Integer position = mapHeaders.get(key);
                convertValue(row, obj.get(key), position);
            }
        }

        for (int i = 0; i < cellNum; i++) {
            sheet.autoSizeColumn(i);
        }

        System.out.println("LOG -> Arquivo xlsx gerado com sucesso!");
        try {
            System.out.println("LOG -> Gravando dados gerados..");
            String newFilename = file.getAbsolutePath().replace(".json", ".xlsx");
            FileOutputStream fileOut = new FileOutputStream(newFilename);
            workbook.write(fileOut);
            fileOut.close();
            System.out.println("LOG -> Arquivo gravado com sucesso em " + newFilename);
        }catch (Exception e){
            System.err.println("ERR -> Erro na gravação do arquivo, motivo: " + e.getMessage());
        }
    }

    private static void convertValue(Row row, Object value, int position){

        if (value instanceof String){
            row.createCell(position).setCellValue((String) value);
        }else if (value instanceof Integer){
            row.createCell(position).setCellValue((Integer) value);
        }else if (value instanceof Boolean) {
            row.createCell(position).setCellValue((Boolean) value);
        }else {
            String s = value.toString();
            if (s.length() > 32000){
                s = s.substring(0, 3200) + "...";
            }
            row.createCell(position).setCellValue(s);
        }
    }

    private static void convertHeader(List<Map<String, Object>> data, Map<String, Integer> mapPosition){

        int count = 0;
        for (Map<String, Object> line : data) {
            for (String chave : line.keySet()) {
                if (!mapPosition.containsKey(chave)){
                    mapPosition.put(chave, count++);
                }
            }
        }
    }

    public static void listf(String directoryName, List<File> files) {

        File directory = new File(directoryName);
        File[] fList = directory.listFiles();

        if(fList == null){
            return;
        }

        for (File file : fList) {
            if (file.isFile() && file.getName().endsWith(".json")) {
                files.add(file);
            }else if (file.isDirectory()) {
                listf(file.getAbsolutePath(), files);
            }
        }
    }
}
