package tianfeng.test;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.TreeMap;

public class ReadExcelFile {

    private static Logger logger = Logger.getLogger(ReadExcelFile.class);

    // 获取工作表信息
    private static int getSheetName(Workbook wb,String name){
        int total = wb.getNumberOfSheets();
        int sheetNu = 0;
        for (int i = 0; i < total; i++){
            System.out.println("第"+(i+1)+"个表格的名字: "+ wb.getSheetName(i));
            if(name.equals(wb.getSheetName(i).trim())){
                sheetNu = i;
            }
        }
        System.out.println("一共有: "+total+"个表格");
        return sheetNu;
    }

    // 操作Excel文件的方法
    private static Map<Integer,Object> getFileObject(String filePath,String name){
        Map<Integer,Object> rowMap = new LinkedHashMap<>();
        try {
            //String encoding = "GBK";
            File excel = new File(filePath);
            if (excel.isFile() && excel.exists()) {   //判断文件是否存在
                String[] split = excel.getName().split("\\.");  //.是特殊字符，需要转义！！！！！
                Workbook wb;
                int sheetNu = 0;
                //根据文件后缀（xls/xlsx）进行判断
                if ( "xls".equals(split[1])){
                    FileInputStream fis = new FileInputStream(excel);  //文件流对象
                    wb = new HSSFWorkbook(fis);
                    // 获取工作表信息
                    sheetNu = getSheetName(wb,name);
                    rowMap = AnalysisData(wb,sheetNu,rowMap);
                }else if ("xlsx".equals(split[1])){
                    wb = new XSSFWorkbook(excel);
                    sheetNu = getSheetName(wb,name);
                    rowMap = AnalysisData(wb,sheetNu,rowMap);
                }else {
                    System.out.println("文件类型错误!");
                }

            } else {
                System.out.println("找不到指定的文件");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return rowMap;
    }

    // 解析文件数据
    public static Map<Integer, Object> AnalysisData(Workbook wb, int sheetNu, Map<Integer, Object> rowMap){
        //开始解析
        Sheet sheet = wb.getSheetAt(sheetNu);     //读取sheet 0

        int firstRowIndex = sheet.getFirstRowNum()+3;   //第一行是列名，所以不读
        int lastRowIndex = sheet.getLastRowNum();
        System.out.println("firstRowIndex: "+firstRowIndex);
        System.out.println("lastRowIndex: "+lastRowIndex);
        logger.info("firstRowIndex: "+firstRowIndex);
        logger.info("lastRowIndex: "+lastRowIndex);

        int num = 0;
        for(int rIndex = firstRowIndex; rIndex <= lastRowIndex; rIndex++) {   //遍历行
            num ++;
            System.out.println("rIndex: " + (rIndex+1));
            logger.info("rIndex: " + (rIndex+1));

            Row row = sheet.getRow(rIndex);
            logger.info("num:"+num+"........................"+row.getCell(2)+"........."+rowMap.size());
            if (row != null) {
                int firstCellIndex = row.getFirstCellNum();
                int lastCellIndex = row.getLastCellNum();
                StringBuffer columnString = new StringBuffer();
                for (int cIndex = firstCellIndex; cIndex < lastCellIndex; cIndex++) {   //遍历列
                    Cell cell = row.getCell(cIndex);
                    //logger.info("单元格属性："+((int)cell.getCellType()));
                    // 通过函数合成的数据无法读取
                    if (cell != null) {
                        switch (cell.getCellType())
                        {
                            case Cell.CELL_TYPE_NUMERIC:
                                DataFormatter dataFormatter = new DataFormatter();
                                dataFormatter.addFormat("###########", null);
                                String phoneNu = dataFormatter.formatCellValue(cell);
//                                        System.out.print(phoneNu+"----");
//                                        logger.info(phoneNu+"----");
                                columnString.append(phoneNu);
                                break;
                            case Cell.CELL_TYPE_STRING:
//                                        System.out.print(cell.getStringCellValue()+"---");
//                                        logger.info(cell.getStringCellValue()+"---");
                                columnString.append(cell.getStringCellValue());
                                break;
                        }
                        if(cIndex < (lastCellIndex-1)){
                            columnString.append("-");
                        }

//                                logger.info(columnString.toString());

                    }else{
                        logger.info("++++++++++++++++++++++");
                    }
                }
                rowMap.put(num,columnString.toString());
                System.out.print(columnString+"\n");
                logger.info(columnString.toString());
            }else{
                logger.info("----------------------------");
            }

        }
        logger.info(num + "---" + rowMap.size()+"+++"+rowMap.keySet().toString());
        return rowMap;
    }



    public static void main(String[] args){

        //excel文件路径
        String excelPath = "F:\\测试\\古城街道已安装999(1).xls";

        // 获取目标文件的数据源
        Map<Integer,Object> map = getFileObject(excelPath,"总表");

        //excel文件路径
        String keyexcelPath = "F:\\测试\\东北片区市民.xls";

        // 获取目标文件的数据源
        Map<Integer,Object> keymap = getFileObject(keyexcelPath,"4组");

        Map< String, Object[] > empinfo = selectData(map,keymap);

        WriteExcelFile.WriteFile(empinfo,keyexcelPath);

    }

    // 数据的筛选
    public static Map< String, Object[] > selectData(Map<Integer,Object> map,Map<Integer,Object> keymap){
        Map< String, Object[] > emp = new TreeMap< String, Object[] >();

        logger.info("-+++++++++++++---数据的筛选----+++++++++++++++++--");
        // 先对总表进行遍历，便于提高效率
        for (Integer index: map.keySet()) {
            String str = (String) map.get(index);
            if (!str.isEmpty()){

                String[] Str = str.split("-");
                if (Str.length >= 5){
                    String name = Str[2];
                    String phone = Str[4];
                    // 先对目标文件
                    for (Integer indx: keymap.keySet()) {
                        String st = (String) keymap.get(indx);
                        if (!st.isEmpty()){
                            String[] keyStr = st.split("-");
                            if (keyStr.length >= 5){
                                String keyname = keyStr[3];
                                String keyphone = keyStr[4];
                                logger.info("-name-"+name+"-keyname----"+keyname);
                                logger.info("-phone-"+phone+"-keyphone----"+keyphone);
                                if (name.equals(keyname)){
                                    logger.info("-name-"+name+"-keyname----"+keyname);
                                    if (phone.equals(keyphone)){
                                        logger.info("-phone-"+phone+"-keyphone----"+keyphone);
                                        emp.put( index.toString(), keyStr);
                                    }
                                }
                            }

                        }
                    }
                }

            }
        }
        return emp;
    }


}
