package tianfeng.test;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jetbrains.annotations.Contract;

import java.io.*;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

public class WriteExcelFile {

    private static Logger logger = Logger.getLogger(WriteExcelFile.class);

    public static void main(String[] args){
        //This data needs to be written (Object[])
        Map< String, Object[] > empinfo =
                new TreeMap< String, Object[] >();
        empinfo.put( "1", new Object[] {
                "EMP ID", "EMP NAME", "DESIGNATION" });
        empinfo.put( "2", new Object[] {
                "tp01", "Gopal", "Technical Manager" });
        empinfo.put( "3", new Object[] {
                "tp02", "Manisha", "Proof Reader" });
        empinfo.put( "4", new Object[] {
                "tp03", "Masthan", "Technical Writer" });
        empinfo.put( "5", new Object[] {
                "tp04", "Satish", "Technical Writer" });
        empinfo.put( "6", new Object[] {
                "tp05", "Krishna", "Technical Writer" });
        empinfo.put( "7", new Object[] {
                "tp05", "Krishna", "Technical Writer" });
        empinfo.put( "67", new Object[] {
                "tp05", "Krishna", "Technical Writer" });
        empinfo.put( "62", new Object[] {
                "tp05", "Krishna", "Technical Writer" });
        //WriteFile(empinfo);
    }
    // 写入文件的方法

    // 创建一个sheet ，并写入指定格式的数据
    @Contract("_, null -> null")
    public static void WriteFile(Map< String, Object[] > empinfo, String path){
        // 打开一个Excel文件
        Workbook wb = OpenFile(path);
        // 创建一个sheet
        // 创建之前查看是否存在
        Sheet spreadsheet = null;
        int flag = getSheetName(wb,"对比结果");
        if (flag == 1){
            wb.removeSheetAt(wb.getSheetIndex("对比结果"));
        }
        spreadsheet = wb.createSheet("对比结果");

        // 准备写入数据
        //Create row object
        Row row;

        // 写入数据
        //Iterate over data and write to sheet
        Set< String > keyid = empinfo.keySet();
        int rowid = 0;
        for (String key : keyid)
        {
            row = spreadsheet.createRow(rowid++);
            Object [] objectArr = empinfo.get(key);
            int cellid = 0;
            for (Object obj : objectArr)
            {
                Cell cell = row.createCell(cellid++);
                cell.setCellValue((String)obj);
            }
        }
        //Write the workbook in file system
        try {
        FileOutputStream out = new FileOutputStream(
                    new File(path));
        wb.write(out);
        out.close();
        System.out.println( path + " written successfully" );
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }


    }

    // 打开一个Excel文件
    public static Workbook OpenFile(String excelPath){
        //excel文件路径
        //String excelPath = "F:\\测试\\东北片区市民.xls";
        try {
            //String encoding = "GBK";
            File excel = new File(excelPath);
            if (excel.isFile() && excel.exists()) {   //判断文件是否存在

                String[] split = excel.getName().split("\\.");  //.是特殊字符，需要转义！！！！！
                Workbook wb;
                //根据文件后缀（xls/xlsx）进行判断
                if ( "xls".equals(split[1])){
                    FileInputStream fis = new FileInputStream(excel);  //文件流对象
                    wb = new HSSFWorkbook(fis);
                    return wb;
                }else if ("xlsx".equals(split[1])){
                    wb = new XSSFWorkbook(excel);
                    return wb;
                }else {
                    System.out.println("文件类型错误!");
                    return null;
                }

            } else {
                System.out.println("找不到指定的文件");
                return null;
            }
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }


    // 判断是否已存在  0-不存在 1-存在
    private static int getSheetName(Workbook wb, String name){
        int total = wb.getNumberOfSheets();
        int sheetNu = 0;
        for (int i = 0; i < total; i++){
            System.out.println("第"+(i+1)+"个表格的名字: "+ wb.getSheetName(i));
            if(name.equals(wb.getSheetName(i).trim())){
                System.out.println("表格已存在！");
                sheetNu = 1;
            }
        }
        System.out.println("一共有: "+total+"个表格");
        return sheetNu;
    }



}
