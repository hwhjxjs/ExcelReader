

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * 读excel
 * @author 钟林森
 *
 */
public class ExcelReader {
    public static void main(String[] args) throws Exception {

        readExcel();

    }

    /**
     * 读取Excel测试，兼容 Excel 2003/2007/2010
     */
    private static String readExcel()
    {
        SimpleDateFormat fmt = new SimpleDateFormat("yyyy-MM-dd");
        ArrayList<ArrayList<String>> allItems = new ArrayList<ArrayList<String>>();
        try {
            //同时支持Excel 2003、2007
            File excelFile = new File("/users/soft/downloads/silvercrest.xlsx"); //创建文件对象
            FileInputStream is = new FileInputStream(excelFile); //文件流
            Workbook workbook = WorkbookFactory.create(is); //这种方式 Excel 2003/2007/2010 都是可以处理的
            int sheetCount = workbook.getNumberOfSheets();  //Sheet的数量
            //遍历每个Sheet
            for (int s = 0; s < sheetCount; s++) {
                Sheet sheet = workbook.getSheetAt(s);
                int rowCount = sheet.getPhysicalNumberOfRows(); //获取总行数



                //遍历每一行
                for (int start = 0; start < 16; start++) {
                    Row row1 = sheet.getRow(1);
                    Cell cell1 = row1.getCell(start);
                    System.out.print(cell1.getStringCellValue());
                    System.out.print("\n");
                    for (int r = 2; r < rowCount; r++) {
                        Row row = sheet.getRow(r);
                        int cellCount = row.getPhysicalNumberOfCells(); //获取总列数
                        ArrayList<String> columns = new ArrayList<String>();
                        //遍历每一列
                        for (int c = 0; c < cellCount; c++) {
                            Cell cell0 = row.getCell(0);
                            Cell cell = row.getCell(c);
                            int cellType = cell.getCellType();
                            String cellValue = cellValue = cell.getStringCellValue();

                            String cellValueNew = "\""+ cell0.getStringCellValue() + "\"" +" = "  + "\"" +cellValue + "\"" +";"+ "\n";
                            if(c == start){
                                columns.add(cellValueNew);
                                System.out.print(cellValueNew);
                            }


                        }

                        allItems.add(columns);
                    }
                    System.out.print("\n\n\n\n");
                }

            }

        }
        catch (Exception e) {
            e.printStackTrace();
        }

        return "success";
    }


}