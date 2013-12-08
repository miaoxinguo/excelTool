package foo.bar.excel.core;

import java.io.BufferedInputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.UnsupportedEncodingException;
import java.util.Properties;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class Main {

    /**
     * 加载properties文件
     */
    private Properties loadProperties(String path) throws IOException {
        InputStream in = new BufferedInputStream(new FileInputStream(path));
        Properties p = new Properties();
        p.load(in);
        return p;
    }

    /**
     * 设置枚举值
     */
    private void setEnmu(Properties p) {
        ColumnIndex.last_month_gy.setIndex(Integer.valueOf(p.getProperty("column_index_gy_last_month")));
        ColumnIndex.last_month_sb.setIndex(Integer.valueOf(p.getProperty("column_index_sb_last_month")));
        ColumnIndex.last_month_6s.setIndex(Integer.valueOf(p.getProperty("column_index_6s_last_month")));
        ColumnIndex.last_month_ldjl.setIndex(Integer.valueOf(p.getProperty("column_index_ldjl_last_month")));
        ColumnIndex.gy_base.setIndex(Integer.valueOf(p.getProperty("column_index_gy_base")));
        ColumnIndex.gy.setIndex(Integer.valueOf(p.getProperty("column_index_gy")));
        ColumnIndex.zl.setIndex(Integer.valueOf(p.getProperty("column_index_zl")));
        ColumnIndex.aq_base.setIndex(Integer.valueOf(p.getProperty("column_index_aq_base")));
        ColumnIndex.sb_base.setIndex(Integer.valueOf(p.getProperty("column_index_sb_base")));
        ColumnIndex.sb.setIndex(Integer.valueOf(p.getProperty("column_index_sb")));
        ColumnIndex.ss_base.setIndex(Integer.valueOf(p.getProperty("column_index_6s_base")));
        ColumnIndex.ss.setIndex(Integer.valueOf(p.getProperty("column_index_6s")));
        ColumnIndex.ldjl_base.setIndex(Integer.valueOf(p.getProperty("column_index_ldjl_base")));
        ColumnIndex.ldjl.setIndex(Integer.valueOf(p.getProperty("column_index_ldjl")));
        ColumnIndex.hlhjy.setIndex(Integer.valueOf(p.getProperty("column_index_hlhjy")));
    }
    
    /**
     * 主方法
     */
    public static void main(String[] args) {
        Main m = new Main();
        Properties p = null;;
        try {
            p = m.loadProperties("setting.properties");
        } catch (IOException e) {
            System.out.println("加载配置文件错误!");
            System.exit(1);
        }
        m.setEnmu(p);
        
        ExcelOperator o = new ExcelOperator();
        HSSFWorkbook targetWorkbook = null;
        try{
            FileInputStream source = new FileInputStream(m.codec(p.getProperty("source_file")));
            FileInputStream template = new FileInputStream(m.codec(p.getProperty("template_file")));
            FileInputStream target = new FileInputStream(m.codec(p.getProperty("target_file")));
            targetWorkbook = o.execute(source, template, target);
        }catch(IOException e){
            e.printStackTrace();
            System.out.println("加载xls文件错误！");
            System.exit(1);
        }catch(Exception e){
            e.printStackTrace();
            System.exit(1);
        }
        
        FileOutputStream fileOut = null;
        try {
            fileOut = new FileOutputStream(m.codec(p.getProperty("target_file")));
            targetWorkbook.write(fileOut);//把Workbook对象输出到 目标文件.xls中  
            fileOut.close();
        } catch (Exception e) {
            System.exit(1);
        }  
        System.out.println("完成!");
    }

    private String codec(String fileName) throws UnsupportedEncodingException{
        return new String(fileName.getBytes("ISO-8859-1"),"utf-8");
    }

}
