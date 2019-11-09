import jxl.Cell;
import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.Colour;
import jxl.format.VerticalAlignment;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Date;


/**
 * @author wangjs6
 * @version 1.0
 * @Description: cOutputFile用于创建结果文件，wOutputFile用于写结果文件
 * @date: 2019/7/10 17:39
 */
public class CWOutputFile {

    // 描述
    private final String PASS = "通过";
    private final String FAIL = "不通过";

    /**
     * @param filepath 文件路径
     * @param caseNo 编号
     * @param testPoint 验证点
     * @param preResult 预期结果
     * @param fresult 实际结果
     * @param testInfo 详细
     * @throws Exception 异常
     */
    public void wOutputFile(String filepath, String caseNo, String testPoint, String preResult, String fresult, String testInfo) throws Exception {
        File output = new File(filepath);
        String result = PASS;
        InputStream instream = new FileInputStream(filepath);
        Workbook readwb = Workbook.getWorkbook(instream);
        // 根据文件创建一个操作对象
        WritableWorkbook wbook = Workbook.createWorkbook(output, readwb);
        WritableSheet readsheet = wbook.getSheet(0);
        // 获取Sheet表中所包含的总行数
        int rsRows = readsheet.getRows(); 

        // 字体样式设置
        WritableFont font = new WritableFont(WritableFont.createFont("宋体"), 10, WritableFont.NO_BOLD);// 字体样式
        WritableCellFormat wcf = new WritableCellFormat(font);

        wcf.setBackground(Colour.BRIGHT_GREEN);

        Cell cell1 = (Cell) readsheet.getCell(0, rsRows);
        if (cell1.getContents().equals("")) {
            if(preResult == null){
                if(fresult == null){
                    preResult = "NULL";
                    fresult = "NULL";

                }else{
                    preResult = "NULL";
                    result = FAIL;
                    wcf.setBackground(Colour.RED);
                }
            }else{
                result = check(preResult,fresult,testInfo);
                if(result.equals(PASS)){
                    wcf.setBackground(Colour.BRIGHT_GREEN);
                }else{
                    wcf.setBackground(Colour.RED);
                }
            }
            // 第1列--案例编号
            Label labetest1 = new Label(0, rsRows, caseNo);
            // 第2列--案例标题
            Label labetest2 = new Label(1, rsRows, testPoint);
            // 第3列--预期值
            Label labetest3 = new Label(2, rsRows, preResult);
            // 第4列--实际值；
            Label labetest4 = new Label(3, rsRows, fresult);
            // 第五列--结果
            Label labetest6 = new Label(4, rsRows, result, wcf);
            // 第6列--详细
            Label labetest5 = new Label(5, rsRows, testInfo);

            readsheet.addCell(labetest1);
            readsheet.addCell(labetest2);
            readsheet.addCell(labetest3);
            readsheet.addCell(labetest4);
            readsheet.addCell(labetest5);
            readsheet.addCell(labetest6);
        }
        wbook.write();
        wbook.close();
    }

    /**
     * 预期值和时间值判断
     * @param hopeStr
     * @param realStr
     * @param testInfo
     * @return
     */
    public String check(String hopeStr, String realStr,String testInfo){
        if(realStr == null){
            return FAIL;
        }
        String result = PASS;
        if(testInfo.contains("数字类型")){
            float a = Float.valueOf(hopeStr);
            float b = Float.valueOf(realStr);
            if(a != b){
                result = FAIL;
            }
        }else if(testInfo.contains("日期类型") || testInfo.contains("时间类型")){
            if(realStr.length() < 10){
                return FAIL;
            }
            if(!hopeStr.equals(realStr.substring(0,4)+realStr.substring(5,7)+realStr.substring(8,10))){
                result = FAIL;
            }
        }else{
            if(!hopeStr.equals(realStr)){
                result = FAIL;
            }
        }
        return result;
    }

    /**
     * cOutputFile方法返回文件路径，作为wOutputFile的入参
     * @param filePath
     * @return
     * @throws IOException
     * @throws WriteException
     */
    public String cOutputFile(String filePath) throws IOException, WriteException {
        // 截取路径用于创建路径 -- D://test/hello  --> D://test/
        String filePathFormat = filePath.substring(0,filePath.lastIndexOf("/"));
        String temp_str = "";
        Date dt = new Date();
        SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmss");
        temp_str = sdf.format(dt);
        File file = new File(filePathFormat);
        if(!file.isDirectory()){
            file.mkdir();
        }
        String filepath = filePath + "output_" + "_" + temp_str + ".xls";
        File output = new File(filepath);

        if (!output.isFile()) {

            output.createNewFile();
            WritableWorkbook writeBook = Workbook.createWorkbook(output);

            WritableSheet Sheet = writeBook.createSheet("输出结果", 0);
            WritableFont headfont = new WritableFont(WritableFont.createFont("宋体"), 11, WritableFont.BOLD);
            WritableCellFormat headwcf = new WritableCellFormat(headfont);

            headwcf.setBackground(Colour.GRAY_25);
            Sheet.setColumnView(0, 11);
            Sheet.setColumnView(1, 40);
            Sheet.setColumnView(2, 11);
            Sheet.setColumnView(3, 11);
            Sheet.setColumnView(4, 18);
            Sheet.setColumnView(5, 50);
            headwcf.setAlignment(Alignment.CENTRE);
            headwcf.setVerticalAlignment(VerticalAlignment.CENTRE);
            Label labe00 = new Label(0, 0, "用例编号", headwcf);
            Label labe10 = new Label(1, 0, "用例标题", headwcf);
            Label labe20 = new Label(2, 0, "预期结果", headwcf);
            Label labe30 = new Label(3, 0, "实际结果", headwcf);
            Label labe40 = new Label(4, 0, "执行结果", headwcf);
            Label labe50 = new Label(5, 0, "执行详细", headwcf);
            Sheet.addCell(labe00);
            Sheet.addCell(labe10);
            Sheet.addCell(labe20);
            Sheet.addCell(labe30);
            Sheet.addCell(labe40);
            Sheet.addCell(labe50);
            writeBook.write();
            writeBook.close();
        }
        return filepath;
    }
}