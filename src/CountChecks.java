import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import java.io.*;
import java.awt.Color;
import org.apache.poi.xssf.model.StylesTable;

public class CountChecks {
	public static void main(String args[]){
		
		XSSFWorkbook report = null;
		FileOutputStream newFile = null;
		
		
		try{
			//Read in excel file
			report = new XSSFWorkbook(new FileInputStream("C:\\Users\\Ryon\\Desktop\\Business_CountReport_06.26.2015.xlsx"));
			
			//Create the style for values that are found to be bad
			Font badFont = report.createFont();
			badFont.setColor(IndexedColors.RED.getIndex());
			CellStyle badCountStyle = report.createCellStyle();
			badCountStyle.setFont(badFont);
			
			//Get the Total Count numberical value from the header for comparison. 
			int totalCount = Integer.parseInt(report.getSheet("Field_Count").getRow(2).getCell(0).getStringCellValue().replaceAll("[\\^A-Za-z:\\s,]", ""));
			
			
			
			
			
//-----------------------------------------------CHECK FIELD_COUNT TAB------------------------------------			
			XSSFSheet fieldCount = report.getSheet("Field_Count");
			for(Row currentRow: fieldCount){
				double currentCountValue;
				double nextCountValue;
				String nextFieldName;
				int rowCount = fieldCount.getLastRowNum();
				XSSFCell currentRowCountCell = (XSSFCell)currentRow.getCell(1);
				XSSFCell nextRowCountCell;
				String currentFieldName = currentRow.getCell(0).getStringCellValue();
				
				if(currentRow.getRowNum() != rowCount)
					nextFieldName = fieldCount.getRow(currentRow.getRowNum() + 1).getCell(0).getStringCellValue();
				
				if(currentRow.getCell(1).getCellType() != Cell.CELL_TYPE_STRING && currentRow.getRowNum() != rowCount && fieldCount.getRow(currentRow.getRowNum() + 1).getCell(1).getCellType() != Cell.CELL_TYPE_STRING){
					currentCountValue = currentRow.getCell(1).getNumericCellValue();
					nextCountValue = fieldCount.getRow(currentRow.getRowNum() + 1).getCell(1).getNumericCellValue();
					nextRowCountCell = (XSSFCell)fieldCount.getRow(currentRow.getRowNum() + 1).getCell(1);
					
					if(currentFieldName.equals("CompanyName") && currentCountValue != totalCount)
						currentRowCountCell.setCellStyle(badCountStyle);
					
					if(currentFieldName.equalsIgnoreCase("Exchange") && currentCountValue != nextCountValue){
						currentRowCountCell.setCellStyle(badCountStyle);
						nextRowCountCell.setCellStyle(badCountStyle);
						
					}
					if(currentFieldName.equals("Year_First") && currentCountValue != totalCount)
						currentRowCountCell.setCellStyle(badCountStyle);
					if(currentFieldName.equals("Year_In_Business_Range") && currentCountValue != totalCount)
						currentRowCountCell.setCellStyle(badCountStyle);
					if(currentFieldName.equals("CreditCode") && currentCountValue != totalCount)
						currentRowCountCell.setCellStyle(badCountStyle);
					if(currentFieldName.equals("Credit_Capacity") && currentCountValue != totalCount)
						currentRowCountCell.setCellStyle(badCountStyle);
					if(currentFieldName.equals("Credit_Description") && currentCountValue != totalCount)
						currentRowCountCell.setCellStyle(badCountStyle);
					if((currentFieldName.equals("UpdateDate") || currentFieldName.equals("AddDate")) && currentCountValue != totalCount)
						currentRowCountCell.setCellStyle(badCountStyle);
					if(currentFieldName.equals("SIC01_Code") && currentCountValue != totalCount)
						currentRowCountCell.setCellStyle(badCountStyle);
					
				
				}//End outer if
				else
					continue;
					
			}//End for
			
			fieldCount = null;
			
//------------------------------------------------CHECK DETAIL_COUNT TAB-----------------------------------
			
			XSSFSheet detailCount = report.getSheet("Detail_Count");
			for(Row currentRow: detailCount){
				double currentCountValue;
				double nextCountValue;
				String nextFieldName;
				int rowCount = detailCount.getLastRowNum();
				XSSFCell currentRowCountCell = (XSSFCell)currentRow.getCell(1);
				XSSFCell nextRowCountCell;
				String currentFieldName = currentRow.getCell(0).getStringCellValue();
				
				if(currentRow.getRowNum() != rowCount)
					nextFieldName = detailCount.getRow(currentRow.getRowNum() + 1).getCell(0).getStringCellValue();
				
				if(currentRow.getCell(1).getCellType() != Cell.CELL_TYPE_STRING && currentRow.getRowNum() != rowCount && detailCount.getRow(currentRow.getRowNum() + 1).getCell(1).getCellType() != Cell.CELL_TYPE_STRING){
					
					
				}//End outer if
			}//End for
			
			
//-----------------------------------------------WRITE CHECKED REPORT--------------------------------------			
			newFile = new FileOutputStream(new File("C:\\Users\\Ryon\\Desktop\\Business_CountReport_06.26.2015_Checked.xlsx"));
			report.write(newFile);
			System.out.println("Written successfully");
		}
		catch(Exception e){
			System.out.println(e.getStackTrace()[0].getLineNumber());
		}
		
		
	}//End main

}//End class
