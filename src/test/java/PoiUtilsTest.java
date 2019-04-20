import com.shawnking07.poiutils.utils.PoiUtils;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;
import org.junit.Test;

public class PoiUtilsTest {
  @Test
  public void readExcel() {
    try {
      final InputStream inputStream = Files.newInputStream(Paths.get("E://aa.xlsx"));
      final List<Model> models = PoiUtils.readExcel(inputStream, Model.class);
      System.out.println(models);
    } catch (IOException | IllegalAccessException e) {
      e.printStackTrace();
    }
  }
}
