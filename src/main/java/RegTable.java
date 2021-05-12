import java.io.Console;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.prefs.Preferences;

/**
 * 往注册表中写数据
 *
 * @author liuyalong
 */
public class RegTable {

    private String keys = "key";
    private String values = "szabsteam";
    private String defaultValue = "not";
    private String node = "/dont_remove";

    public String getKeys() {
        return keys;
    }

    public void setKeys(String keys) {
        this.keys = keys;
    }

    public String getValues() {
        return values;
    }

    public void setValues(String values) {
        this.values = values;
    }

    public String getDefaultValue() {
        return defaultValue;
    }

    public void setDefaultValue(String defaultValue) {
        this.defaultValue = defaultValue;
    }

    public String getNode() {
        return node;
    }

    public void setNode(String node) {
        this.node = node;
    }

    public RegTable(int endTime) {
        boolean b = this.checkTime(endTime);
        if (!b) {
            System.out.println("有效期不足!");
            System.exit(1);
        }
    }

    /**
     * @param endTime 检查当前时间是否超过预设时间,
     * @return 没超过预设时间为true
     */
    public boolean checkTime(int endTime) {

        SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd");
        int now = Integer.parseInt(sdf.format(new Date()));
        return (endTime - now) >= 0;
    }

    // 把相应的值储存到变量中去
    public void writeValue() {
        // HKEY_CURRENT_USER\Software\JavaSoft\prefs下写入注册表值.
        Preferences pre = Preferences.userNodeForPackage(this.getClass()).node(node);
        pre.put(keys, values);
    }

    /***
     * 根据key获取value
     *
     */
    public String getValue(String key) {
        Preferences pre = Preferences.userNodeForPackage(this.getClass()).node(node);
        return pre.get(key, defaultValue);
    }

    public void checkValue() {
        String key = this.getValue("key");
        while (true) {
            if (!key.equals(this.values)) {
                System.out.println("请输入密码...");
                Console console = System.console();
                String s = new String(console.readPassword());
                if (!s.equals(this.values)) {
                    System.out.println("密码错误!\n");
                } else {
                    this.writeValue();
                    break;
                }
            } else {
                break;
            }
        }
    }

}
