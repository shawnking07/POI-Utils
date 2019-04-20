package com.shawnking07.poiutils.utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

public class CellUtils {
  public static boolean isEmpty(Cell cell) {
    if (cell == null) {
      return true;
    }
    if (cell.getCellType().equals(CellType.BLANK)) {
      return true;
    }
    return cell.getCellType().equals(CellType.STRING) && "".equals(cell.getStringCellValue());
  }

  public static String getString(Cell cell) {
    if (cell == null) {
      throw new NullPointerException("cell cannot be null");
    }
    if (cell.getCellType().equals(CellType.STRING)) {
      return cell.getRichStringCellValue().getString().trim();
    } else if (cell.getCellType().equals(CellType.NUMERIC)) {
      return String.valueOf(cell.getNumericCellValue());
    }
    return "";
  }
}
