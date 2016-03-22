package com.vlad;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.Scanner;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

public class shop {
	
	public static void main(String[] args) throws IOException { 
		System.out.println("Здравствуйте! Добро пожаловать в наш магазин!");
		int menu = 1; 
		while (menu != 5){ 
		System.out.println(); 
		System.out.println("1. Продажа товара \n2. Поставка товара \n3. Данные по продажам \n4. Удаление со склада\n5. Выйти из магазина");  
		Scanner in = new Scanner(System.in); 
		menu = in.nextInt(); 
		FileInputStream sklad1 = new FileInputStream("C:/Users/Владислав/Desktop/sklad.xls"); 
		Workbook wb_sklad = new HSSFWorkbook(sklad1); 
		FileInputStream sales1 = new FileInputStream("C:/Users/Владислав/Desktop/sales.xls"); 
		Workbook wb_sales = new HSSFWorkbook(sales1); 
		BufferedReader br = new BufferedReader (new InputStreamReader(System.in)); 
		switch (menu) 
		{ 
		   case 1: 
		      out (wb_sklad);
		      int vvod = 1; 
		      while (vvod != 3) 
		      {  
		    	  
		         double quantity_sklad;
		         int i=-1; 
		         int id_sales = 0; 
		         int stroka_sales = 0; 
		         System.out.println("1. Ввод по названию \n2. Ввод по id \n3. Запрос продажи"); 
		         vvod = in.nextInt(); 
		         switch (vvod) 
		         { 
		             case 1: 
		                System.out.println("Введите имя и колличество"); 
		                String name_sklad = br.readLine(); 
		                quantity_sklad = in.nextDouble();
		                sale (name_sklad, quantity_sklad, i, id_sales, stroka_sales, wb_sklad, wb_sales);
		             break; 
		             case 2: 
		                System.out.println("Введите id и колличество"); 
		                String id_sklad = br.readLine(); 
		                quantity_sklad = in.nextDouble();
		                sale (id_sklad, quantity_sklad, i, id_sales, stroka_sales, wb_sklad, wb_sales);
		             break; 
		             case 3: 
		                FileOutputStream sales2 = new FileOutputStream("C:/Users/Владислав/Desktop/sales.xls"); 
		                wb_sales.write(sales2); 
		                sales2.close(); 
		                FileOutputStream sklad2 = new FileOutputStream("C:/Users/Владислав/Desktop/sklad.xls"); 
		                wb_sklad.write(sklad2); 
		                sklad2.close(); 
		             break; 
		          } 
		       } 
		    break; 
		    case 2: 
		         int stroka = 0; 
		         for (org.apache.poi.ss.usermodel.Row row: wb_sklad.getSheetAt(0)) 
		           {
		        	   stroka++; 
		               for (Cell cell: row);
		           } 

		         System.out.println("Введите id и количество"); 
		         String id_sklad = br.readLine(); 
		         int j=-1; 
		         int proverka = 1; 
		         double quantity_postavka = in.nextDouble(); 
		         for (org.apache.poi.ss.usermodel.Row row: wb_sklad.getSheetAt(0)) 
		           { 
		              j++; 
		              for (Cell cell: row) 
		                 if (getCellText(cell).equals(id_sklad)) 
		                 { 
		                	 proverka=0; 
		                     cell = wb_sklad.getSheetAt(0).getRow(j).getCell(3); 
		                     cell.setCellValue(cell.getNumericCellValue()+quantity_postavka); 
		                     FileOutputStream sklad2 = new FileOutputStream("C:/Users/Владислав/Desktop/sklad.xls"); 
		                     wb_sklad.write(sklad2); 
		                     sklad2.close(); 
		                 } 
		           } 

		         if (proverka == 1) 
		           { 
		              System.out.println("Введите название и цену"); 
		              String newname = br.readLine(); 
		              double price = in.nextDouble(); 
		              Row row = wb_sklad.getSheetAt(0).createRow(stroka); 

		              Cell cell1 = row.createCell(0); 
		              cell1.setCellValue(id_sklad); 
		              Cell cell2 = row.createCell(1); 
		              cell2.setCellValue(newname); 
		              Cell cell3 = row.createCell(2); 
		              cell3.setCellValue(price); 
		              Cell cell4 = row.createCell(3); 
		              cell4.setCellValue(quantity_postavka); 
		              FileOutputStream sklad2 = new FileOutputStream("C:/Users/Владислав/Desktop/sklad.xls"); 
		              wb_sklad.write(sklad2); 
		              sklad2.close(); 
		           } 
		         out (wb_sklad);
		    break; 
		    case 3: 
		         int j2=-1; 
		         out (wb_sales);
		         System.out.println("Введите id продажи"); 
		         int id_sales = in.nextInt(); 

		         for (org.apache.poi.ss.usermodel.Row row: wb_sales.getSheetAt(1)) 
		         {
		        	 j2++; 
		             for (Cell cell: row) 
		             { 
		                if (j2==id_sales) 
		                { 
		                   System.out.print(getCellText(cell)); 
		                   System.out.print(" "); 
		                } 
		             } 
		         } 
		    break; 
		    case 4:
		    	int i2=-1;
		    	System.out.println("Введите пароль");
		    	int parol = in.nextInt();
		    	if (parol == 1234)
		    	{
		    	System.out.println("Введите id товара, который надо удалить из склада");
		    	String id_sklad_delete = br.readLine();
		    	w:
		    	for (org.apache.poi.ss.usermodel.Row row: wb_sklad.getSheetAt(0)) 
		         { 
		          i2++; 
		          for (Cell cell: row)
		          {
		        	  if (getCellText(cell).equals(id_sklad_delete)) 
		        	  {
		        		 removeRow(i2, wb_sklad);
		        		 break w;
		        	  } 
		          }
		         }
		    	
		        		  
		    	 out(wb_sklad);
		    	 FileOutputStream sklad3 = new FileOutputStream("C:/Users/Владислав/Desktop/sklad.xls"); 
		         wb_sklad.write(sklad3); 
		         sklad3.close(); 
		    	}
		    	else System.out.println("Пароль неверный");
		    break;
		} 

		sales1.close();
		sklad1.close();
		}
		System.out.println("До свидания! Приходите ещё!");
		} 


		//функции
		public static void removeRow(int rowIndex, Workbook wb) 
		{ 
		   int lastRowNum = wb.getSheetAt(0).getLastRowNum(); 
		   if(rowIndex >= 0 && rowIndex < lastRowNum)
		   { 
		      wb.getSheetAt(0).shiftRows(rowIndex+1,lastRowNum, -1); 
		   } 
		   if(rowIndex == lastRowNum)
		   { 
		      Row removingRow = wb.getSheetAt(0).getRow(rowIndex); 
		      if(removingRow != null)
		      {  
		    	 wb.getSheetAt(0).removeRow(removingRow); 
		      } 
		   } 
		}


		public static String getCellText(Cell cell)
		{ 
		   String result=""; 
		   switch (cell.getCellType()) 
		   { 
		      case Cell.CELL_TYPE_STRING: 
		         result = cell.getRichStringCellValue().getString(); 
		      break; 
		      case Cell.CELL_TYPE_NUMERIC: 
		    	 result = Double.toString(cell.getNumericCellValue());
		      break; 
		   } 
		return result; 
		} 

		public static void sale (String id_name_sklad, double quantity_sklad, int i, int id_sales, int stroka_sales, Workbook wb1, Workbook wb2)
		{
			   for (org.apache.poi.ss.usermodel.Row row: wb1.getSheetAt(0)) 
		       { 
		        i++; 
		        for (Cell cell: row) 
		           if (getCellText(cell).equals(id_name_sklad)) 
		           {   
		               cell = wb1.getSheetAt(0).getRow(i).getCell(3); 
		               if (cell.getNumericCellValue() < quantity_sklad) 
		                   System.out.println("Товара не хватает на складе"); 
		               else 
		               { 
		                   for (org.apache.poi.ss.usermodel.Row row2: wb2.getSheetAt(0)) 
		                   {
		                	   stroka_sales++; 
		                       for (Cell cell2: row2);
		                   } 

		                   Row row_sales1 = wb2.getSheetAt(0).createRow(stroka_sales);
		                   id_sales = stroka_sales; 
		                   Cell cell2 = row_sales1.createCell(0); 
		                   cell2.setCellValue(Integer.toString(id_sales)); 
		                   Cell cell3 = wb1.getSheetAt(0).getRow(i).getCell(2); 
		                   Cell cell4 = row_sales1.createCell(1); 
		                   cell4.setCellValue(quantity_sklad*cell3.getNumericCellValue()); 
		                   
		                   cell.setCellValue(cell.getNumericCellValue()-quantity_sklad); 

		                   cell = wb1.getSheetAt(0).getRow(i).getCell(1); 
		                   Row row_sales2 = wb2.getSheetAt(1).createRow(stroka_sales); 
		                   Cell cell5 = row_sales2.createCell(0); 
		                   cell5.setCellValue(getCellText(cell)); 
		                   Cell cell6 = row_sales2.createCell(1); 
		                   cell6.setCellValue(quantity_sklad); 
		               } 
		           }
		        } 
		}

		public static void out(Workbook wb1)
		{
			 for (org.apache.poi.ss.usermodel.Row row: wb1.getSheetAt(0))
		     { 
		        for (Cell cell: row)
		        { 
		           System.out.print(getCellText(cell)); 
		           System.out.print(" "); 
		        } 
		          
		        System.out.println(); 
		     } 
		}

}
