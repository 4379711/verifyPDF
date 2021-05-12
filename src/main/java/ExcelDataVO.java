import lombok.Data;
import lombok.ToString;

import java.io.Serializable;

/**
 * @author liuyalong
 * @date 2020/9/18 14:18
 */
@Data
@ToString
public class ExcelDataVO implements Serializable {
    private String fileName;
    private String signDate;
    private String validBefore;
    private String validAfter;
    private String subject;
    private String serialNumber;
    private Boolean isEffective;
    private String filePath;
}
