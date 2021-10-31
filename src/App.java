import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class App {
    public static void main(String[] args) throws Exception {
        
        String filename2 = "../Secondary_filter-_lowscores.xls";

        String filename3 = "../secondary_gRNAs.xls";
        
        List<List<HSSFCell>> sheet1Data = getExcelData(filename3);
        List<List<HSSFCell>> sheetData = getExcelData(filename2);

        //parse_files_upstream(sheet1Data, sheetData);

        parse_files_a50_1F(sheet1Data, sheetData);

    }


    private static void parse_files(List<List<HSSFCell>> sheet1Data, List<List<HSSFCell>> sheetData) throws Exception
    {
        ArrayList<ArrayList<String>> generationData = extractGenerationData(sheetData);
        ArrayList<ArrayList<String>> grnaData = extractGRNAData(sheet1Data);
        ArrayList<ArrayList<String>> geneGRNA = new ArrayList<ArrayList<String>>();
        ArrayList<String> GRNArow =  new ArrayList<String>();
        String start  = "";
        String end = "";
        String name = "";
        String Gstart = "";
        String Gend = "";
        String score = "";

        System.out.println("GRNA score 0: ");

        for (ArrayList<String> gene: generationData)
        {
            geneGRNA = new ArrayList<ArrayList<String>>();

            start  = gene.get(0);
            end = gene.get(1);
            name = gene.get(2);
            for ( ArrayList<String> row : grnaData)
            {  
                Gstart = row.get(1);
                Gend = row.get(2);
                score = row.get(3);
                
                if (Integer.valueOf(Gstart) >= Integer.valueOf(start) && Integer.valueOf(Gend) <=Integer.valueOf(end) && Integer.valueOf(score) == 0) //start-end-score
                {
                    GRNArow =  new ArrayList<String>();
                    for (String column : row)
                    {
                        GRNArow.add(column);
                    }
                    geneGRNA.add(GRNArow);
                }
                

            }

            
            if(geneGRNA.size() > 0)
            {
                createExcel_0(name, geneGRNA);
            }
            
            /*
            System.out.println(name);
            System.out.println("");
            for (ArrayList<String> row: geneGRNA)
            {
                
                for(String column : row)
                {
                    System.out.print(column+" | ");
                }
                System.out.println("");
                System.out.println("");
            }
            */
        
            


        }

        System.out.println("GRNA score >=50: ");

        for (ArrayList<String> gene: generationData)
        {
            geneGRNA = new ArrayList<ArrayList<String>>();

            start  = gene.get(0);
            end = gene.get(1);
            name = gene.get(2);
            for ( ArrayList<String> row : grnaData)
            {  
                Gstart = row.get(1);
                Gend = row.get(2);
                score = row.get(3);
                
                if (Integer.valueOf(Gstart) >= Integer.valueOf(start) && Integer.valueOf(Gend) <=Integer.valueOf(end) && Integer.valueOf(score) >= 50)//start-end-score
                {
                    GRNArow =  new ArrayList<String>();
                    for (String column : row)
                    {
                        GRNArow.add(column);
                    }
                    geneGRNA.add(GRNArow);
                }
                

            }

            if(geneGRNA.size() > 0)
            {
                createExcel_a50(name, geneGRNA);
            }
            

            System.out.println(name);
            System.out.println("");
            for (ArrayList<String> row: geneGRNA)
            {
                
                for(String column : row)
                {
                    System.out.print(column+" | ");
                }
                System.out.println("");
                System.out.println("");
            }
        
            


        }

        System.out.println("GRNA score <50: ");

        for (ArrayList<String> gene: generationData)
        {
            geneGRNA = new ArrayList<ArrayList<String>>();

            start  = gene.get(0);
            end = gene.get(1);
            name = gene.get(2);
            for ( ArrayList<String> row : grnaData)
            {  
                Gstart = row.get(1);
                Gend = row.get(2);
                score = row.get(3);
                
                if (Integer.valueOf(Gstart) >= Integer.valueOf(start) && Integer.valueOf(Gend) <=Integer.valueOf(end) && Integer.valueOf(score) < 50)//start-end-score
                {
                    GRNArow =  new ArrayList<String>();
                    for (String column : row)
                    {
                        GRNArow.add(column);
                    }
                    geneGRNA.add(GRNArow);
                }
                

            }

            
            if(geneGRNA.size() > 0)
            {
                createExcel_b50(name, geneGRNA);
            }
            

            System.out.println(name);
            System.out.println("");
            for (ArrayList<String> row: geneGRNA)
            {
                
                for(String column : row)
                {
                    System.out.print(column+" | ");
                }
                System.out.println("");
                System.out.println("");
            }
        
            


        }
    }

    
    private static void parse_files_upstream(List<List<HSSFCell>> sheet1Data, List<List<HSSFCell>> sheetData) throws Exception
    {
        ArrayList<ArrayList<String>> generationData = extractGenerationData(sheetData);
        ArrayList<ArrayList<String>> generationDataCorrected = extractGenerationData2(sheetData);
        ArrayList<ArrayList<String>> grnaData = extractGRNAData(sheet1Data);
        ArrayList<ArrayList<String>> geneGRNA = new ArrayList<ArrayList<String>>();
        ArrayList<String> GRNArow =  new ArrayList<String>();
        String start  = "";
        String end = "";
        String name = "";
        String Gstart = "";
        String Gend = "";
        String score = "";
        geneGRNA = new ArrayList<ArrayList<String>>();
        for (ArrayList<String> gene: generationData)
        {
            

            start  = gene.get(0);
            end = gene.get(1);
            name = gene.get(2);
            //System.out.println(name);
            String startCorrected = "";
            String endCorrected = "";
            String nameCorrected = "";
            
            for ( ArrayList<String> row : grnaData)
            {  
                Gstart = row.get(1);
                Gend = row.get(2);
                score = row.get(3);
                boolean overlap = false;
                for(ArrayList<String> geneCorrected: generationDataCorrected)
                {
                    startCorrected = geneCorrected.get(0);
                    endCorrected = geneCorrected.get(1);
                    nameCorrected = geneCorrected.get(2);

                    if(nameCorrected != name && Integer.valueOf(Gstart) >= Integer.valueOf(startCorrected) && Integer.valueOf(Gend) <= Integer.valueOf(endCorrected))
                    {
                        overlap = true;
                    }
                }
                
                if (overlap == false && Integer.valueOf(Gstart) >= Integer.valueOf(start) && Integer.valueOf(Gend) <=Integer.valueOf(end) && Integer.valueOf(score) >=0)//start-end-score
                {
                    GRNArow =  new ArrayList<String>();
                    GRNArow.add(name);
                    for (String column : row)
                    {
                        GRNArow.add(column);
                    }
                    geneGRNA.add(GRNArow);
                    //System.out.println("No 0verlap on gene: "+name);
                }
                

            }
        }

        createExcel_a50_1F(geneGRNA);

    }

    private static void parse_files_a50_1F(List<List<HSSFCell>> sheet1Data, List<List<HSSFCell>> sheetData) throws Exception
    {
        ArrayList<ArrayList<String>> generationData = extractGenerationData(sheetData);
        ArrayList<ArrayList<String>> grnaData = extractGRNAData(sheet1Data);
        ArrayList<ArrayList<String>> geneGRNA = new ArrayList<ArrayList<String>>();
        ArrayList<String> GRNArow =  new ArrayList<String>();
        String start  = "";
        String end = "";
        String name = "";
        String Gstart = "";
        String Gend = "";
        String score = "";
        geneGRNA = new ArrayList<ArrayList<String>>();
        for (ArrayList<String> gene: generationData)
        {
            

            start  = gene.get(1);
            end = gene.get(2);
            name = gene.get(0);
            for ( ArrayList<String> row : grnaData)
            {  
                Gstart = row.get(1);
                Gend = row.get(2);
                score = row.get(3);
                
                if (Integer.valueOf(Gstart) >= Integer.valueOf(start) && Integer.valueOf(Gend) <=Integer.valueOf(end)  )//start-end-score
                {
                    GRNArow =  new ArrayList<String>();
                    GRNArow.add(name);
                    for (String column : row)
                    {
                        GRNArow.add(column);
                    }
                    geneGRNA.add(GRNArow);
                }
                

            }
            
            /*
            System.out.println(name);
            System.out.println("");
            for (ArrayList<String> row: geneGRNA)
            {
                
                for(String column : row)
                {
                    System.out.print(column+" | ");
                }
                System.out.println("");
                System.out.println("");
            }
            */
        }

        createExcel_a50_1F(geneGRNA);

    }

    private static void createExcel_a50_1F( ArrayList<ArrayList<String>> list) throws Exception
    {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Sheet1");

        int rowCount = -1;
        for (ArrayList<String> data : list)
        {
            Row row = sheet.createRow(++rowCount);

            int columnCount = -1;

            for (String data2 : data)
            {
                Cell cell = row.createCell(++columnCount);
                if (data2 instanceof String) {
                    cell.setCellValue((String) data2);
                }
            }

        }
        try (FileOutputStream outputStream = new FileOutputStream("output/"+"Secondary.xlsx")) //output file name
        {
            workbook.write(outputStream);
        }
    }

    private static void createExcel_0(String name, ArrayList<ArrayList<String>> list) throws Exception
    {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Sheet1");

        int rowCount = -1;
        for (ArrayList<String> data : list)
        {
            Row row = sheet.createRow(++rowCount);

            int columnCount = -1;

            for (String data2 : data)
            {
                Cell cell = row.createCell(++columnCount);
                if (data2 instanceof String) {
                    cell.setCellValue((String) data2);
                }
            }

        }
        try (FileOutputStream outputStream = new FileOutputStream("output/0/"+name+".xlsx")) //output file name
        {
            workbook.write(outputStream);
        }
    }

    private static void createExcel_a50(String name, ArrayList<ArrayList<String>> list) throws Exception
    {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Sheet1");

        int rowCount = -1;
        for (ArrayList<String> data : list)
        {
            Row row = sheet.createRow(++rowCount);

            int columnCount = -1;

            for (String data2 : data)
            {
                Cell cell = row.createCell(++columnCount);
                if (data2 instanceof String) {
                    cell.setCellValue((String) data2);
                }
            }

        }
        try (FileOutputStream outputStream = new FileOutputStream("output/above_50/"+name+".xlsx")) //output file name
        {
            workbook.write(outputStream);
        }
    }

    private static void createExcel_b50(String name, ArrayList<ArrayList<String>> list) throws Exception
    {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Sheet1");

        int rowCount = -1;
        for (ArrayList<String> data : list)
        {
            Row row = sheet.createRow(++rowCount);

            int columnCount = -1;

            for (String data2 : data)
            {
                Cell cell = row.createCell(++columnCount);
                if (data2 instanceof String) {
                    cell.setCellValue((String) data2);
                }
            }

        }
        try (FileOutputStream outputStream = new FileOutputStream("output/below_50/"+name+".xlsx")) //output file name
        {
            workbook.write(outputStream);
        }
    }


    private static List<List<HSSFCell>> getExcelData(String filename) {
        List<List<HSSFCell>> sheetData = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(filename)) {
            HSSFWorkbook workbook = new HSSFWorkbook(fis);
            HSSFSheet sheet = workbook.getSheetAt(0);
            Iterator rows = sheet.rowIterator();
            while (rows.hasNext()) {
                HSSFRow row = (HSSFRow) rows.next();
                Iterator cells = row.cellIterator();

                List<HSSFCell> data = new ArrayList<>();
                while (cells.hasNext()) {
                    HSSFCell cell = (HSSFCell) cells.next();
                    data.add(cell);
                }
                sheetData.add(data);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return sheetData;
    }

    

    private static ArrayList<ArrayList<String>> extractGenerationData(List<List<HSSFCell>> sheetData) {
        int column = 0;
        boolean skip = true;
        ArrayList<ArrayList<String>> table = new ArrayList<ArrayList<String>>();
        ArrayList<String> row  = new ArrayList<String>();
        DataFormatter df = new DataFormatter();
        for (List<HSSFCell> data : sheetData) {
            column = 0;
            for (HSSFCell cell : data) {
                boolean finished = false;                
                if(column == 2)
                {
                    
                    String value = df.formatCellValue(cell);
                    String start ="";
                    String end = "";
                    String tmp = "";
                    for(int i = 0; i<value.length();i++)
                    {
                        
                        if (value.charAt(i)==':')
                        {
                            tmp = "";
                            i++;
                        }
                        else if (value.charAt(i)=='-' && !finished)
                        {
                            start = tmp ;
                            row.add(start);
                            i++;
                            tmp = "";
                            finished = true;
                        }
                        else if (value.charAt(i)==' ')
                        {
                            end = tmp ;
                            row.add(end);
                        }
                        tmp+=value.charAt(i);
                    }

                    
                   
                    
                }else if(column == 0)
                {
                    
                    String value = df.formatCellValue(cell);
                    row.add(value);
                    
                }
                column++;
                
                    
            }
            if (skip == false)
            {
                table.add(row);
            }
            else
            {
                skip = false;
            }
            row = new ArrayList<String>();
        }
        return table;
    }

    private static ArrayList<ArrayList<String>> extractGenerationData2(List<List<HSSFCell>> sheetData) {
        int column = 0;
        boolean skip = true;
        ArrayList<ArrayList<String>> table = new ArrayList<ArrayList<String>>();
        ArrayList<String> row  = new ArrayList<String>();
        DataFormatter df = new DataFormatter();
        for (List<HSSFCell> data : sheetData) {
            column = 0;
            for (HSSFCell cell : data) {
                
                if(column == 24 || column == 25 || column == 26)
                {
                    
                    String value = df.formatCellValue(cell);
                    row.add(value);
                    
                }
                column++;
                
                    
            }
                if (skip == false)
                {
                    table.add(row);
                }
                else
                {
                    skip = false;
                }
                row = new ArrayList<String>();
        }
        return table;
    }


    private static ArrayList<ArrayList<String>> extractGRNAData(List<List<HSSFCell>> sheetData) {
        int column = 0;
        ArrayList<ArrayList<String>> table = new ArrayList<ArrayList<String>>();
        ArrayList<String> row  = new ArrayList<String>();
        for (List<HSSFCell> data : sheetData) {
            column = 0;
            for (HSSFCell cell : data) {

                if (column == 11)
                {
                    
                    DataFormatter df = new DataFormatter();
                    String value = df.formatCellValue(cell);
                    if (!value.contains("TTTT"))
                    {
                        row.add(value);
                    }
                
                }else
                {
                    DataFormatter df = new DataFormatter();
                    String value = df.formatCellValue(cell);

                    row.add(value);
                }
                column++;
                
                    
            }
            if (row.size()== 19)
            {
                table.add(row);
            }
                
                row = new ArrayList<String>();
        }
        return table;
    }
}