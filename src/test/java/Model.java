import com.google.common.base.Objects;
import com.shawnking07.poiutils.annotation.Column;

public class Model {
  @Column("name")
  private String name;

  @Column("times")
  private Integer times;

  @Override
  public boolean equals(Object o) {
    if (this == o) {
      return true;
    }
    if (o == null || getClass() != o.getClass()) {
      return false;
    }
    Model model = (Model) o;
    return Objects.equal(name, model.name) &&
        Objects.equal(times, model.times);
  }

  @Override
  public int hashCode() {
    return Objects.hashCode(name, times);
  }

  @Override
  public String toString() {
    return "Model{" +
        "name='" + name + '\'' +
        ", times=" + times +
        '}';
  }
}
