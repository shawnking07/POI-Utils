package com.shawnking07.poiutils.utils;

import com.google.common.collect.ImmutableBiMap;
import com.google.common.collect.ImmutableBiMap.Builder;
import com.shawnking07.poiutils.annotation.Column;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PoiUtils {
  public static <T> List<T> readExcel(InputStream inputStream, Class<T> tClass)
      throws IOException, IllegalAccessException {
    final XSSFWorkbook wb = new XSSFWorkbook(inputStream);
    final XSSFSheet sheet = wb.getSheetAt(0);
    final ImmutableBiMap<Integer, Field> index = getColumnIndex(sheet, tClass).inverse();
    final ArrayList<T> ts = new ArrayList<>();
    for (int i = 1; i < sheet.getLastRowNum(); i++) {
      final T t;
      try {
        t = tClass.getDeclaredConstructor().newInstance();
      } catch (InstantiationException | IllegalAccessException | InvocationTargetException | NoSuchMethodException e) {
        throw new IllegalArgumentException("model class error");
      }
      final XSSFRow row = sheet.getRow(i);
      for (int j = 0; j < row.getLastCellNum(); j++) {
        final Field field = index.get(j);
        if (field == null) {
          continue;
        }
        field.setAccessible(true);
        field.set(t, convert(field.getType(), row.getCell(j)));
      }
      ts.add(t);
    }
    return ts;
  }

  private static Object convert(Class fieldType, Cell value) {
    if (String.class.equals(fieldType)) {
      return CellUtils.getString(value);
    } else if (Integer.class.equals(fieldType)) {
      return Double.valueOf(CellUtils.getString(value)).intValue();
    } else if (Double.class.equals(fieldType)) {
      return Double.valueOf(CellUtils.getString(value));
    }
    return null;
  }

  private static <T> ImmutableBiMap<Field, Integer> getColumnIndex(XSSFSheet sheet, Class<T> tClass) {
    final ImmutableBiMap<String, Field> columnField = getFieldColumn(tClass).inverse();
    final XSSFRow row = sheet.getRow(0);
    final Builder<Field, Integer> builder = ImmutableBiMap.builder();
    for (int i = 0; i < row.getLastCellNum(); i++) {
      final String header = CellUtils.getString(row.getCell(i));
      if (columnField.keySet().contains(header)) {
        builder.put(columnField.get(header), i);
      }
    }
    final ImmutableBiMap<Field, Integer> build = builder.build();
    if (build.size() < columnField.size()) {
      throw new IllegalArgumentException("template error");
    }
    return build;
  }

  /**
   * column header map
   *
   * @param tClass
   * @param <T>
   * @return java field value -> column value
   */
  private static <T> ImmutableBiMap<Field, String> getFieldColumn(Class<T> tClass) {
    final Field[] declaredFields = tClass.getDeclaredFields();
    final Map<Field, String> collect =
        Arrays.stream(declaredFields)
            .filter(v -> v.isAnnotationPresent(Column.class))
            .collect(Collectors.toMap(v -> v, v -> v.getAnnotation(Column.class).value()));
    return ImmutableBiMap.<Field, String>builder().putAll(collect).build();
  }
}
