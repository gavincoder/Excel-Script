package excelToolPro;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelTools {
		
		/**
		 * description:读取xlsx文件指定的若干列单元格数据
		 */
		@SuppressWarnings({ "resource", "unused" })
		public ArrayList<ArrayList<String>> xlsx_reader(String excel_url,int ... args) throws IOException {
	        ArrayList<ArrayList<String>> ans=new ArrayList<ArrayList<String>>();
	        File excelFile = null;
	        InputStream is = null;
	        try {
	        	 //读取xlsx文件
		        XSSFWorkbook xssfWorkbook = null;
		        //寻找目录读取文件
		        excelFile = new File(excel_url);
		        is = new FileInputStream(excelFile);
		        xssfWorkbook = new XSSFWorkbook(is);		      
		        if(xssfWorkbook==null){
		        	System.out.println("未读取到内容,请检查路径！");
		        	return null;
		        }		        
		        //遍历xlsx中的sheet
		        for (int numSheet = 0; numSheet < xssfWorkbook.getNumberOfSheets(); numSheet++) {
		            XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(numSheet);
		            if (xssfSheet == null) {
		                continue;
		            }
		            // 对于每个sheet，读取其中的每一行
		            for (int rowNum = 0; rowNum <= xssfSheet.getLastRowNum(); rowNum++) {
		                XSSFRow xssfRow = xssfSheet.getRow(rowNum);
		                if (xssfRow == null) continue; 
		                ArrayList<String> curarr=new ArrayList<String>();
		                for(int columnNum = 0 ; columnNum<args.length ; columnNum++){
		                	XSSFCell cell = xssfRow.getCell(args[columnNum]);		                	
		                	curarr.add( Trim_str( getValue(cell) ) );
		                }
		                ans.add(curarr);
		            }
		        }
	        }catch(Exception e) {
	        	e.getStackTrace();
	        }finally {
	        	  is.close();
	        }
	        return ans;
	    }
	    
	    //判断后缀为xlsx的excel文件的数据类型
	    @SuppressWarnings("deprecation")
		private static String getValue(XSSFCell xssfRow) {
	    	if(xssfRow==null){
	    		return "---";
	    	}
	        if (xssfRow.getCellType() == xssfRow.CELL_TYPE_BOOLEAN) {
	            return String.valueOf(xssfRow.getBooleanCellValue());
	        } else if (xssfRow.getCellType() == xssfRow.CELL_TYPE_NUMERIC) {
	        	double cur=xssfRow.getNumericCellValue();
	        	long longVal = Math.round(cur);
	        	Object inputValue = null;
	        	if(Double.parseDouble(longVal + ".0") == cur)  
	        	        inputValue = longVal;  
	        	else  
	        	        inputValue = cur; 
	            return String.valueOf(inputValue);
	        } else if(xssfRow.getCellType() == xssfRow.CELL_TYPE_BLANK || xssfRow.getCellType() == xssfRow.CELL_TYPE_ERROR){
	        	return "---";
	        }
	        else {
	            return String.valueOf(xssfRow.getStringCellValue());
	        }
	    }
	    
	    //判断后缀为xls的excel文件的数据类型
	    @SuppressWarnings("deprecation")
		private static String getValue(HSSFCell hssfCell) {
	    	if(hssfCell==null){
	    		return "---";
	    	}
	        if (hssfCell.getCellType() == hssfCell.CELL_TYPE_BOOLEAN) {
	            return String.valueOf(hssfCell.getBooleanCellValue());
	        } else if (hssfCell.getCellType() == hssfCell.CELL_TYPE_NUMERIC) {
	        	double cur=hssfCell.getNumericCellValue();
	        	long longVal = Math.round(cur);
	        	Object inputValue = null;
	        	if(Double.parseDouble(longVal + ".0") == cur)  
	        	        inputValue = longVal;  
	        	else  
	        	        inputValue = cur; 
	            return String.valueOf(inputValue);
	        } else if(hssfCell.getCellType() == hssfCell.CELL_TYPE_BLANK || hssfCell.getCellType() == hssfCell.CELL_TYPE_ERROR){
	        	return "---";
	        }
	        else {
	            return String.valueOf(hssfCell.getStringCellValue());
	        }
	    }
	    
	  //字符串修剪  去除所有空白符号 ， 问号 ， 中文空格
	    static private String Trim_str(String str){
	        if(str==null)
	            return null;
	        return str.trim();//str.replaceAll("[\\s\\?]", "").replace("　", "");
	    }
	    
	    /**
		 * description:批量修改指定sheet表中某列的多行单元格格式（颜色）
	     */
	    public  void setXLSXColor(String url,int sheetNumber,ArrayList<Integer> rowNumbers,int colNumber) throws IOException {//XLSX 的单元格填充颜色
        	 //读取xlsx文件
	        InputStream is = new FileInputStream(new File(url));	    	
	    	XSSFWorkbook my_workbook = new XSSFWorkbook(is);    	
            XSSFCellStyle my_style = my_workbook.createCellStyle(); // Get access to XSSFCellStyle */
            my_style.setFillPattern(XSSFCellStyle.FINE_DOTS );
            my_style.setFillForegroundColor(IndexedColors.BRIGHT_GREEN.getIndex());
            //my_style.setFillBackgroundColor(IndexedColors.RED.getIndex());
            is.close();      
            XSSFSheet xssfSheet = my_workbook.getSheetAt(sheetNumber);  //拿取xlsx中的sheetNumber,设置单元格颜色
            XSSFCell cellTarge = null;
            for(int i=0;i < rowNumbers.size();i++) {
            	cellTarge = xssfSheet.getRow(rowNumbers.get(i)).getCell(colNumber);
            	cellTarge.setCellStyle(my_style);
            }     
            FileOutputStream out = new FileOutputStream(new File(url));
            my_workbook.write(out);
            out.close();
	    }
}
