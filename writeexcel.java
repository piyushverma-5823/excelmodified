package test;
import  java.io.*;  
import java.util.*;
import  org.apache.poi.hssf.usermodel.HSSFSheet;  
import  org.apache.poi.hssf.usermodel.HSSFWorkbook;  
import  org.apache.poi.hssf.usermodel.HSSFRow;  
public class writeexcel  
{  
public static void main(String[] args)   
{  
try   
{  
//declare file name to be create   
String filename = "C:\\Users\\piyush.verma\\Desktop\\StudentsDetail.xlsx";   
//creating an instance of HSSFWorkbook class  
HSSFWorkbook workbook = new HSSFWorkbook();  
//invoking creatSheet() method and passing the name of the sheet to be created   
HSSFSheet sheet = workbook.createSheet("Basic Details");   
//creating the 0th row using the createRow() method  
HSSFRow rowhead = sheet.createRow((short)0);  
//creating cell by using the createCell() method and setting the values to the cell by using the setCellValue() method  
rowhead.createCell(0).setCellValue("S.No.");  
rowhead.createCell(1).setCellValue("Student Name");  
rowhead.createCell(2).setCellValue("Roll Number");  
rowhead.createCell(3).setCellValue("e-mail");  
rowhead.createCell(4).setCellValue("Current Percentage");  
//creating the row  
for(int i=0;i<2;i++)
{
	Scanner sc=new Scanner(System.in);
HSSFRow row = sheet.createRow((short)(i+1));  
//inserting data in the first row  
row.createCell(0).setCellValue(sc.next());  
row.createCell(1).setCellValue(sc.next());  
row.createCell(2).setCellValue(sc.next());  
row.createCell(3).setCellValue(sc.next());  
row.createCell(4).setCellValue(sc.next());  
}

FileOutputStream fileOut = new FileOutputStream(filename);  
workbook.write(fileOut);  
//closing the Stream  
fileOut.close();  
//closing the workbook  
workbook.close();  
//prints the message on the console  
System.out.println("Excel file has been written successfully.");  
}   
catch (Exception FileNotFound)   
{  
FileNotFound.printStackTrace();  
}  
}
}  