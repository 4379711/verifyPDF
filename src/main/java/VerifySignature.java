import com.spire.pdf.PdfDocument;
import com.spire.pdf.security.PdfCertificate;
import com.spire.pdf.security.PdfSignature;
import com.spire.pdf.widget.PdfFormFieldWidgetCollection;
import com.spire.pdf.widget.PdfFormWidget;
import com.spire.pdf.widget.PdfSignatureFieldWidget;

import javax.swing.*;
import java.awt.*;
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;


/**
 * @author liuyalong
 */
public class VerifySignature {
    private static String fromDirPath;
    private static String toFilePath;

    public static void main(String[] args) {
//        RegTable regTable = new RegTable(20300101);
//        regTable.checkValue();

        final JFrame jf = new JFrame("校验电子印章签名");
        jf.setResizable(false);
        jf.setVisible(true);
        jf.setSize(400, 280);
        jf.setLocationRelativeTo(null);
        jf.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);

        JPanel panel = new JPanel();

        // 创建文本区域, 用于显示相关信息
        final JTextArea msgTextArea = new JTextArea(10, 30);

        msgTextArea.setLineWrap(true);
        panel.add(msgTextArea);

        JButton openBtn = new JButton("选择文件路径");
        openBtn.addActionListener(e -> showFileOpenDialog(jf, msgTextArea));
        panel.add(openBtn);

        JButton saveBtn = new JButton("结果保存位置");
        saveBtn.addActionListener(e -> showFileSaveDialog(jf, msgTextArea));
        panel.add(saveBtn);

        jf.setContentPane(panel);
        jf.setVisible(true);

        JButton enSureBtn = new JButton("确认");
        enSureBtn.addActionListener(e -> enSureListener(jf));
        panel.add(enSureBtn);

        jf.setContentPane(panel);
        jf.setVisible(true);
    }

    /*
     * 打开文件
     */
    private static void showFileOpenDialog(Component parent, JTextArea msgTextArea) {
        // 创建一个默认的文件选取器
        JFileChooser fileChooser = new JFileChooser();

        // 设置默认显示的文件夹为当前文件夹
        fileChooser.setCurrentDirectory(new File("."));

        // 设置文件选择的模式（只选文件、只选文件夹、文件和文件均可选）
        fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

        // 设置是否允许多选
        fileChooser.setMultiSelectionEnabled(false);

//        // 添加可用的文件过滤器（FileNameExtensionFilter 的第一个参数是描述, 后面是需要过滤的文件扩展名 可变参数）
//        fileChooser.addChoosableFileFilter(new FileNameExtensionFilter("zip(*.zip, *.rar)", "zip", "rar"));
//
//        // 设置默认使用的文件过滤器
//        fileChooser.setFileFilter(new FileNameExtensionFilter("image(*.jpg, *.png, *.gif)", "jpg", "png", "gif"));

        // 打开文件选择框（线程将被阻塞, 直到选择框被关闭）
        int result = fileChooser.showOpenDialog(parent);

        if (result == JFileChooser.APPROVE_OPTION) {
            // 如果点击了"确定", 则获取选择的文件路径
            File file = fileChooser.getSelectedFile();
            fromDirPath = file.getAbsolutePath();

            msgTextArea.append("选择源文件: " + fromDirPath + "\n\n");
        }
    }

    /*
     * 选择文件保存路径
     */
    private static void showFileSaveDialog(Component parent, JTextArea msgTextArea) {
        // 创建一个默认的文件选取器
        JFileChooser fileChooser = new JFileChooser();

        //把时间戳经过处理得到期望格式的时间
        Date date = new Date();
        SimpleDateFormat format0 = new SimpleDateFormat("yyyyMMddHHmmss");
        String now = format0.format(date.getTime());

        // 设置打开文件选择框后默认输入的文件名
        fileChooser.setSelectedFile(new File(now + ".xlsx"));

        // 打开文件选择框（线程将被阻塞, 直到选择框被关闭）
        int result = fileChooser.showSaveDialog(parent);

        if (result == JFileChooser.APPROVE_OPTION) {
            // 如果点击了"保存", 则获取选择的保存路径
            File file = fileChooser.getSelectedFile();
            toFilePath = file.getAbsolutePath();
            msgTextArea.append("结果文件路径: " + toFilePath + "\n\n");
        }
    }

    //找到需要的内容
    public final static Pattern PATTERN = Pattern.compile("\\[Subject\\].*?CN=(.*?),.*?\\[Issuer\\](.*?)\\[Serial Number\\](.*?)\\[Not Before\\](.*?)\\[Not After\\](.*?)\\[Thumbprint\\](.*?)");
    // 剔除特殊字符
    public final static Pattern REPLACE_PATTERN = Pattern.compile("\t|\r|\n");

    /**
     * 查找某个路径下的所有pdf文件
     *
     * @return 所有的pdf绝对路径
     */
    public static HashSet<String> listDir(String path) {
        HashSet<String> FileNameString = new HashSet<String>();

        if (path == null) {
            return null;
        }
        //获取其file对象
        File file = new File(path);
        //遍历path下的文件和目录，放在File数组中
        File[] fs = file.listFiles();
        if (fs == null) {
            System.out.println(path + "路径下没有文件");
            return null;
        }

        //遍历File[]数组
        for (File f : fs) {
            String fileName = String.valueOf(f);
            //若非目录(即文件)，则打印
            if (!f.isDirectory() && fileName.toLowerCase().endsWith(".pdf")) {
                FileNameString.add(fileName);
            }
        }
        return FileNameString;
    }

    /**
     * 检验pdf文件是否签名
     */
    public static ArrayList<ExcelDataVO> checkPdf(String filePath) {
        //创建PdfDocument实例
        PdfDocument doc = new PdfDocument();

        //创建结果集
        ArrayList<ExcelDataVO> arrayList = new ArrayList<>();

        //文件名,注意windows下应该是\\,linux下是/
        String fileName = filePath.substring(filePath.lastIndexOf("\\") + 1);

        //加载含有签名的PDF文件
        doc.loadFromFile(filePath);

        //获取域集合
        PdfFormWidget pdfFormWidget = (PdfFormWidget) doc.getForm();
        PdfFormFieldWidgetCollection pdfFormFieldWidgetCollection = pdfFormWidget.getFieldsWidget();
        int count = pdfFormFieldWidgetCollection.getCount();
//        System.out.println("签名域个数:" + count);


        //遍历域
        for (int i = 0; i < count; i++) {
            //判定是否为签名域
            if (pdfFormFieldWidgetCollection.get(i) instanceof PdfSignatureFieldWidget) {
                //获取签名域
                PdfSignatureFieldWidget signatureFieldWidget = (PdfSignatureFieldWidget) pdfFormFieldWidgetCollection.get(i);

                ExcelDataVO excelDataVO = new ExcelDataVO();
                //设置文件绝对路径
                excelDataVO.setFilePath(filePath);
                excelDataVO.setFileName(fileName);

                //获取签名时间
                PdfSignature signature = signatureFieldWidget.getSignature();
                excelDataVO.setSignDate(String.valueOf(signature.getDate()));

                //获取签名的内容
                PdfCertificate certificate = signature.getCertificate();

//                System.out.println(certificate.get_IssuerName().getName());

//                excelDataVO.setSubject(String.valueOf(certificate.getSubject()));

                String certificateString = certificate.toString();

                Matcher m = REPLACE_PATTERN.matcher(certificateString);
                certificateString = m.replaceAll("");

                Matcher matcher = PATTERN.matcher(certificateString);

                //判定签名是否有效
                boolean result = signature.verifySignature();

                //文件是否被更改
                boolean resultModified = signature.verifyDocModified();

                while (matcher.find()) {
//                    String group = matcher.group(0);
                    String subject = matcher.group(1);

//                    String issuer = matcher.group(2);
                    String serialNumber = matcher.group(3);

                    String before = matcher.group(4);
                    String after = matcher.group(5);
//                    String sha1 = matcher.group(6);
                    excelDataVO.setSubject(subject);
                    excelDataVO.setSerialNumber(serialNumber);
                    excelDataVO.setValidBefore(before);
                    excelDataVO.setValidAfter(after);
                    //因为盖章有顺序,后盖的章因为改动了文件本书,导致先盖的章验证失败,所有这里要判断是否被修改过
                    excelDataVO.setIsEffective(result || resultModified);

                }
                arrayList.add(excelDataVO);
            }
        }
        return arrayList;
    }

    /*
     * 开始执行业务逻辑
     */
    private static void enSureListener(JFrame parent) {
        parent.dispose();
        System.out.println("开始验签...");

        //从某个路径下获取所有的pdf文件路径
        HashSet<String> filePaths = listDir(fromDirPath);
        if (filePaths == null) {
            System.out.println("验签失败,请选择文件...");
            return;
        }
        List<ExcelDataVO> excelDataVOS = new ArrayList<>();

        for (String filePath : filePaths) {
            ArrayList<ExcelDataVO> excelDataVOS1 = checkPdf(filePath);
            excelDataVOS.addAll(excelDataVOS1);
        }

        ExcelWriter.writeExcel(excelDataVOS, toFilePath);
        System.out.println("验签完成...");
    }
}
