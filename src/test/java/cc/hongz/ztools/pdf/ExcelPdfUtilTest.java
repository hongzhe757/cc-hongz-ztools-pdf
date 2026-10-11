package cc.hongz.ztools.pdf;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

@DisplayName("Excel转PDF工具类测试")
public class ExcelPdfUtilTest {

    /**
     * 测试用临时目录。
     */
    @TempDir
    private File tempDir;

    @Test
    @DisplayName("含中文的xlsx转PDF生成有效文件测试")
    public void excelToPdfTest() throws IOException {
        final File excel = new File(this.tempDir, "source.xlsx");
        writeWorkbook(excel, 1);
        final File pdf = new File(this.tempDir, "target.pdf");

        ExcelPdfUtil.excelToPdf(excel, pdf);

        assertValidPdf(pdf);
    }

    @Test
    @DisplayName("合并单元格与样式还原测试")
    public void excelToPdfWithMergedCellTest() throws IOException {
        final File excel = new File(this.tempDir, "styled.xlsx");
        writeStyledWorkbook(excel);
        final File pdf = new File(this.tempDir, "styled.pdf");

        // 合并单元格、背景色、粗体、列宽均应正常处理而不抛异常
        ExcelPdfUtil.excelToPdf(excel, pdf);

        assertValidPdf(pdf);
    }

    @Test
    @DisplayName("多工作表全部渲染测试")
    public void excelToPdfWithMultipleSheetsTest() throws IOException {
        final File singleExcel = new File(this.tempDir, "single.xlsx");
        writeWorkbook(singleExcel, 1);
        final File singlePdf = new File(this.tempDir, "single.pdf");
        ExcelPdfUtil.excelToPdf(singleExcel, singlePdf);

        final File multiExcel = new File(this.tempDir, "multi.xlsx");
        writeWorkbook(multiExcel, 3);
        final File multiPdf = new File(this.tempDir, "multi.pdf");
        ExcelPdfUtil.excelToPdf(multiExcel, multiPdf);

        assertValidPdf(singlePdf);
        assertValidPdf(multiPdf);
        // 三个工作表应渲染出比单个工作表更多的内容
        assertTrue(multiPdf.length() > singlePdf.length(), "多工作表产出的 PDF 应更大");
    }

    /**
     * 生成含中文内容的多工作表 xlsx。
     *
     * @param file       目标文件
     * @param sheetCount 工作表数量
     * @throws IOException 写入失败时抛出
     */
    private static void writeWorkbook(final File file, final int sheetCount) throws IOException {
        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream out = new FileOutputStream(file)) {
            for (int i = 0; i < sheetCount; i++) {
                final Sheet sheet = workbook.createSheet("工作表" + (i + 1));
                final Row header = sheet.createRow(0);
                header.createCell(0).setCellValue("中文表头");
                header.createCell(1).setCellValue("列二");
                header.createCell(2).setCellValue("列三");

                final Row data = sheet.createRow(1);
                data.createCell(0).setCellValue("第一行");
                data.createCell(1).setCellValue(123.45D);
                data.createCell(2).setCellValue("尾部");
            }
            workbook.write(out);
        }
    }

    /**
     * 生成含合并单元格、背景色与粗体样式的 xlsx。
     *
     * @param file 目标文件
     * @throws IOException 写入失败时抛出
     */
    private static void writeStyledWorkbook(final File file) throws IOException {
        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream out = new FileOutputStream(file)) {
            final Sheet sheet = workbook.createSheet("样式表");

            final CellStyle titleStyle = workbook.createCellStyle();
            titleStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            final Font boldFont = workbook.createFont();
            boldFont.setBold(true);
            boldFont.setFontHeightInPoints((short) 14);
            titleStyle.setFont(boldFont);

            final Row titleRow = sheet.createRow(0);
            final org.apache.poi.ss.usermodel.Cell titleCell = titleRow.createCell(0);
            titleCell.setCellValue("跨三列的标题");
            titleCell.setCellStyle(titleStyle);
            // 非左上角单元格的内容应被忽略
            titleRow.createCell(1).setCellValue("应被忽略");
            titleRow.createCell(2).setCellValue("应被忽略");
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 2));

            final Row dataRow = sheet.createRow(1);
            dataRow.createCell(0).setCellValue("纵向合并");
            dataRow.createCell(1).setCellValue("值一");
            dataRow.createCell(2).setCellValue("值二");
            final Row thirdRow = sheet.createRow(2);
            thirdRow.createCell(1).setCellValue("值三");
            thirdRow.createCell(2).setCellValue("值四");
            sheet.addMergedRegion(new CellRangeAddress(1, 2, 0, 0));

            sheet.setColumnWidth(0, 30 * 256);
            sheet.setColumnWidth(1, 20 * 256);
            sheet.setColumnWidth(2, 20 * 256);
            workbook.write(out);
        }
    }

    /**
     * 断言文件是一个有效的 PDF。
     *
     * @param pdf 待校验的 PDF 文件
     * @throws IOException 读取失败时抛出
     */
    private static void assertValidPdf(final File pdf) throws IOException {
        assertTrue(pdf.isFile(), "PDF 文件应已生成");
        assertTrue(pdf.length() > 0, "PDF 文件不应为空");
        final byte[] bytes = Files.readAllBytes(pdf.toPath());
        final String header = new String(bytes, 0, 5, StandardCharsets.ISO_8859_1);
        assertEquals("%PDF-", header, "PDF 文件头应为 %PDF-");
    }
}
