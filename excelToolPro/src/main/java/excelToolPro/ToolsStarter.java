package excelToolPro;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.Set;


public class ToolsStarter {
	public static void main(String[] args) throws IOException  {
		ExcelTools test= new ExcelTools(); 
		ArrayList<ArrayList<String>> arrA=test.xlsx_reader("D:/excelTest/A.xlsx",1);  //后面的参数代表需要输出哪些列，参数个数可以任意
		ArrayList<ArrayList<String>> arrB=test.xlsx_reader("D:/excelTest/B.xlsx",1);//已翻译字段集合
		Set<String> setB = new HashSet<>();
		ArrayList<Integer> colorRowNumbers = new ArrayList<>();
		
		for(int i=0;i<arrB.size();i++){//行数
			ArrayList<String> row=arrB.get(i);
			for(int j=0;j<row.size();j++){//列数
				setB.add(row.get(j));
			}
		}
		for(int i=0;i<arrA.size();i++){//行数
			ArrayList<String> row = arrA.get(i);
			for(int j=0;j<row.size();j++){//列数
				if(!"---".equals(row.get(j)) && !"".equals(row.get(j).trim())&& setB.contains(row.get(j))) {
					colorRowNumbers.add(i);
					System.out.print(i+"---" + row.get(j)+",");//打印重复的行数
				}
			}
			//System.out.print(",");
		}
		test.setXLSXColor("D:/excelTest/A.xlsx",0,colorRowNumbers,1);//文件，sheet号，行号，列号
		System.out.println(" set color done");
  
	}
}

/*//打印输出
for(int i=0;i<arr.size();i++){//行数
	ArrayList<String> row=arr.get(i);
	for(int j=0;j<row.size();j++){//列数
		System.out.print(row.get(j)+" ");//某行某列数值
	}
	System.out.println("");//一行一行输出
}*/