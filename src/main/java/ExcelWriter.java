import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.ss.usermodel.Font;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * @author liuyalong
 * @date 2020/9/18 14:20
 */
public class ExcelWriter {

    //表头
    private static final List<String> CELL_HEADS;

    static {
        // 类装载时就载入指定好的表头信息，如有需要，可以考虑做成动态生成的表头
        CELL_HEADS = new ArrayList<>();
        CELL_HEADS.add("文件");
        CELL_HEADS.add("签名时间");
        CELL_HEADS.add("有效起始时间");
        CELL_HEADS.add("有效截至时间");
        CELL_HEADS.add("签名者");
        CELL_HEADS.add("序列号");
        CELL_HEADS.add("是否通过验签");
        CELL_HEADS.add("文件地址");
    }

    /**
     * 生成Excel并写入数据信息
     *
     * @param dataList 数据列表
     * @return 写入数据后的工作簿对象
     */
    public static Workbook exportData(List<ExcelDataVO> dataList) {
        // 生成xlsx的Excel
        Workbook workbook = new SXSSFWorkbook();

        // 如需生成xls的Excel，请使用下面的工作簿对象，注意后续输出时文件后缀名也需更改为xls
        //Workbook workbook = new HSSFWorkbook();

        // 生成Sheet表，写入第一行的表头
        Sheet sheet = buildDataSheet(workbook);
        //构建每行的数据内容
        int rowNum = 1;
        for (ExcelDataVO data : dataList) {
            if (data == null) {
                continue;
            }
            //输出行数据
            Row row = sheet.createRow(rowNum++);
            convertDataToRow(workbook, data, row);
        }

        return workbook;
    }

    /**
     * 生成sheet表，并写入第一行数据（表头）
     *
     * @param workbook 工作簿对象
     * @return 已经写入表头的Sheet
     */
    private static Sheet buildDataSheet(Workbook workbook) {
        Sheet sheet = workbook.createSheet();
        // 设置表头宽度
        for (int i = 0; i < CELL_HEADS.size(); i++) {
            sheet.setColumnWidth(i, 4000);
        }
        // 设置默认行高
        sheet.setDefaultRowHeight((short) 400);
        // 构建头单元格样式
        CellStyle cellStyle = buildHeadCellStyle(sheet.getWorkbook());
        // 写入第一行各列的数据
        Row head = sheet.createRow(0);
        for (int i = 0; i < CELL_HEADS.size(); i++) {
            Cell cell = head.createCell(i);
            cell.setCellValue(CELL_HEADS.get(i));
            cell.setCellStyle(cellStyle);
        }
        return sheet;
    }

    /**
     * 设置第一行表头的样式
     *
     * @param workbook 工作簿对象
     * @return 单元格样式对象
     */
    private static CellStyle buildHeadCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        //对齐方式设置
        style.setAlignment(HorizontalAlignment.CENTER);
        //边框颜色和宽度设置
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); // 下边框
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); // 左边框
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex()); // 右边框
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex()); // 上边框
        //设置背景颜色
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        //粗体字设置
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        return style;
    }


    /**
     * 将数据转换成行
     *
     * @param data 源数据
     * @param row  行对象
     */
    private static void convertDataToRow(Workbook workbook, ExcelDataVO data, Row row) {


        int cellNum = 0;
        Cell cell;

        //对特殊数值设置颜色
        CellStyle cellStyle = workbook.createCellStyle();

        //字体设置
        Font font = workbook.createFont();
        font.setBold(true);
        font.setColor(IndexedColors.GREEN.getIndex());
        cellStyle.setFont(font);

        // 文件名
        cell = row.createCell(cellNum++);
        cell.setCellValue(data.getFileName());

        // 签名时间
        cell = row.createCell(cellNum++);
        cell.setCellValue(null == data.getSignDate() ? "" : data.getSignDate());

        // 有效期
        cell = row.createCell(cellNum++);
        cell.setCellValue(null == data.getValidBefore() ? "" : data.getValidBefore());

        // 有效期
        cell = row.createCell(cellNum++);
        cell.setCellValue(null == data.getValidAfter() ? "" : data.getValidAfter());
        //主题
        cell = row.createCell(cellNum++);
        cell.setCellValue(null == data.getSubject() ? "" : data.getSubject());
        //序列号
        cell = row.createCell(cellNum++);
        cell.setCellValue(null == data.getSerialNumber() ? "" : data.getSerialNumber());
        //是否通过验签
        cell = row.createCell(cellNum++);
        if (data.getIsEffective()) {
            cell.setCellValue("签名有效");
        } else {
            cell.setCellValue("签名无效");
            cell.setCellStyle(cellStyle);
        }
        //文件路径
        cell = row.createCell(cellNum);
        Hyperlink hyperlink = workbook.getCreationHelper().createHyperlink(HyperlinkType.FILE);
        File pdfFilePath = new File(data.getFilePath());

        String relativeURI = pdfFilePath.toURI().toString();
        hyperlink.setAddress(relativeURI);
        cell.setHyperlink(hyperlink);
        cell.setCellStyle(cellStyle);
    }

    public static void writeExcel(List<ExcelDataVO> dataVOList, String exportFilePath) {

        // 写入数据到工作簿对象内
        Workbook workbook = ExcelWriter.exportData(dataVOList);

        // 以文件的形式输出工作簿对象
        FileOutputStream fileOut = null;
        try {
            File exportFile = new File(exportFilePath);
            if (!exportFile.exists()) {
                boolean newFile = exportFile.createNewFile();
                if (!newFile) {
                    System.out.println("文件创建失败");
                }
            }

            fileOut = new FileOutputStream(exportFilePath);
            workbook.write(fileOut);
            fileOut.flush();
        } catch (Exception e) {
            System.out.println("输出Excel时发生错误，错误原因：" + e.getMessage());
        } finally {
            try {
                if (null != fileOut) {
                    fileOut.close();
                }
                workbook.close();
            } catch (IOException e) {
                System.out.println("关闭输出流时发生错误，错误原因：" + e.getMessage());
            }
        }
    }
}
