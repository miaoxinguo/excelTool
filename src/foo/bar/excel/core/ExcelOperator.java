package foo.bar.excel.core;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Row;

/**
 * Excel操作
 *
 * @author miaoxinguo2002@gmail.com
 * @version base 2013-12-6
 */
public class ExcelOperator {
    
    /**
     * 自动填数据
     */
    public HSSFWorkbook execute(FileInputStream source, FileInputStream template, FileInputStream target){
        HSSFWorkbook targetWorkbook = load(target);
        HSSFSheet sourceSheet = load(source).getSheetAt(0);
        HSSFSheet templateSheet = load(template).getSheetAt(0);
        
        // 模板文件、上月文件、目标文件的 行迭代器
        Iterator<Row> templateRowIterator = templateSheet.rowIterator();
        Iterator<Row> sourceRowIterator = sourceSheet.rowIterator();
        Iterator<Row> targetRowIterator = targetWorkbook.getSheetAt(0).rowIterator();

        // 为提高遍历效率 存入集合
        List<Row> sourceRowList = toList(sourceRowIterator);
        List<Row> targetRowList = toList(targetRowIterator);
        
        // 遍历模板中的所有记录
        while(templateRowIterator.hasNext()){
            Row templateRow = templateRowIterator.next();
            // 忽略第一行
            if(templateRow.getRowNum()==0) {
                continue;  
            }
            String name = templateRow.getCell(0).getStringCellValue();    // 姓名
            
            // 根据姓名从源文件中获得所在行
            Row sourceRow = null;
            for(int i=0; i<sourceRowList.size(); i++){
                Row row = sourceRowList.get(i);
                if(name.equals(row.getCell(2).getStringCellValue())){
                    sourceRow = row;
                    sourceRowList.remove(i);
                    break;
                }
            }
            
            // 将模板文件中的数据按姓名设置到目标文件中
            for(int i=0; i<targetRowList.size(); i++){
                Row targetRow = targetRowList.get(i);
                // targetRow rourceRow 已经过处理 直接用即可
                if(name.equals(targetRow.getCell(2).getStringCellValue())){
                    // 设置各项加减分的值
                    targetRow.getCell(ColumnIndex.zl.getIndex()).setCellValue(templateRow.getCell(1).getNumericCellValue());  // 质量
                    targetRow.getCell(ColumnIndex.gy.getIndex()).setCellValue(templateRow.getCell(2).getNumericCellValue());  // 工艺
                    targetRow.getCell(ColumnIndex.aq.getIndex()).setCellValue(templateRow.getCell(3).getNumericCellValue());  // 安全
                    targetRow.getCell(ColumnIndex.sb.getIndex()).setCellValue(templateRow.getCell(4).getNumericCellValue());  // 设备
                    targetRow.getCell(ColumnIndex.ss.getIndex()).setCellValue(templateRow.getCell(5).getNumericCellValue());  // 6s
                    targetRow.getCell(ColumnIndex.ldjl.getIndex()).setCellValue(templateRow.getCell(6).getNumericCellValue());  // 劳动纪律
                    targetRow.getCell(ColumnIndex.hlhjy.getIndex()).setCellValue(templateRow.getCell(7).getNumericCellValue());  // 合理化建议
                    
                    // 设置各项基础分
                    if(sourceRow != null){
                        double gy = sourceRow.getCell(ColumnIndex.last_month_gy.getIndex()).getNumericCellValue();
                        targetRow.getCell(ColumnIndex.gy_base.getIndex()).setCellValue(gy);  // 工艺
                        double aq = sourceRow.getCell(ColumnIndex.last_month_aq.getIndex()).getNumericCellValue();
                        targetRow.getCell(ColumnIndex.aq_base.getIndex()).setCellValue(aq);  // 安全
                        double sb = sourceRow.getCell(ColumnIndex.last_month_sb.getIndex()).getNumericCellValue();
                        targetRow.getCell(ColumnIndex.sb_base.getIndex()).setCellValue(sb);  // 设备
                        double ss = sourceRow.getCell(ColumnIndex.last_month_6s.getIndex()).getNumericCellValue();
                        targetRow.getCell(ColumnIndex.ss_base.getIndex()).setCellValue(ss);  // 6s
                        double ldjl = sourceRow.getCell(ColumnIndex.last_month_ldjl.getIndex()).getNumericCellValue();
                        targetRow.getCell(ColumnIndex.ldjl_base.getIndex()).setCellValue(ldjl);  // 劳动记录
                    } else{
                        System.out.println("上月文件中未找到姓名为"+name+"的记录，基础分数无法设置");
                    }
                    targetRowList.remove(i);
                    break;
                }
            }
        }
        // 如果targetRowList 剩余元素大于0 说明目标文件的条目数要比模板中的多
        if(targetRowList.size() > 0){
            System.out.println("目标文件中"+targetRowList.size()+"条记录未被设置,分别是：");
            for(Row row : targetRowList){
                System.out.println(row.getRowNum()+" "+row.getCell(2).getStringCellValue());
            }
        }
        // 执行表中的所有公式
        HSSFFormulaEvaluator.evaluateAllFormulaCells(targetWorkbook);
        return targetWorkbook;
    }

    /**
     * 迭代器中有效 行存到list中
     * @return 
     */
    private List<Row> toList(Iterator<Row> rowIterator) {
        List<Row> rowList = new ArrayList<>();
        while(rowIterator.hasNext()){
            Row row = rowIterator.next();
            // 行号小于2、列数小于2的行（防止末尾多一行）、姓名为空的行 不添加到list中
            if(row.getRowNum()<2 || row.getPhysicalNumberOfCells() < 3 || row.getCell(2).getStringCellValue()==null){
                continue;
            }
            rowList.add(row);
        }
        return rowList;
    }
    
    /**
     * 从jar包同级路径中加载模板 和 目标文件
     */
    public HSSFWorkbook load(FileInputStream in) {
        POIFSFileSystem file = null;
        try {
            file = new POIFSFileSystem(in);
            return new HSSFWorkbook(file);
        } catch (IOException e) {
            throw new RuntimeException();
        }
    }
}
