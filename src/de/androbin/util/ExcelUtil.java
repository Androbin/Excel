package de.androbin.util;

import org.apache.poi.ss.usermodel.*;

public final class ExcelUtil
{
	private ExcelUtil()
	{
	}
	
	public static Object getCellValue( final Cell cell )
	{
		switch ( cell.getCellType() )
		{
			case Cell.CELL_TYPE_BOOLEAN :
				return cell.getBooleanCellValue();
			
			case Cell.CELL_TYPE_ERROR :
				return cell.getErrorCellValue();
			
			case Cell.CELL_TYPE_FORMULA :
				return cell.getCellFormula();
			
			case Cell.CELL_TYPE_NUMERIC :
				return cell.getNumericCellValue();
			
			case Cell.CELL_TYPE_STRING :
				return cell.getStringCellValue();
		}
		
		return null;
	}
	
	public static Object getCellValueOrDefault( final Cell cell, final Object o )
	{
		final Object value = getCellValue( cell );
		return value == null ? o : value;
	}
}