package cc.hongz.ztools.pdf;

import cc.hongz.ztools.core.exceptions.UtilException;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.junit.jupiter.api.Assumptions;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

@DisplayName("Word转PDF工具类测试")
public class WordPdfUtilTest {

    /**
     * 测试用临时目录。
     */
    @TempDir
    private File tempDir;

    @Test
    @DisplayName("docx转PDF生成有效PDF文件测试")
    public void wordToPdfTest() throws IOException {
        final File docx = new File(this.tempDir, "source.docx");
        writeChineseDocx(docx);
        final File pdf = new File(this.tempDir, "target.pdf");

        WordPdfUtil.wordToPdf(docx, pdf);

        assertValidPdf(pdf);
    }

    @Test
    @DisplayName("指定中文字体转PDF测试")
    public void wordToPdfWithFontTest() throws IOException {
        final String fontPath = WordPdfUtil.detectChineseFont();
        Assumptions.assumeTrue(fontPath != null, "当前环境未探测到中文字体，跳过");

        final File docx = new File(this.tempDir, "font-source.docx");
        writeChineseDocx(docx);
        final File pdf = new File(this.tempDir, "font-target.pdf");

        WordPdfUtil.wordToPdf(docx, pdf, fontPath);

        assertValidPdf(pdf);
    }

    @Test
    @DisplayName("字体文件不存在时抛出异常测试")
    public void invalidFontTest() throws IOException {
        final File docx = new File(this.tempDir, "invalid-font-source.docx");
        writeChineseDocx(docx);
        final File pdf = new File(this.tempDir, "invalid-font-target.pdf");
        final String notExistFont = new File(this.tempDir, "not-exist.ttf").getAbsolutePath();

        assertThrows(UtilException.class, () -> WordPdfUtil.wordToPdf(docx, pdf, notExistFont));
    }

    @Test
    @DisplayName("中文字体探测结果合法测试")
    public void detectChineseFontTest() {
        final String fontPath = WordPdfUtil.detectChineseFont();

        if (fontPath != null) {
            assertTrue(new File(fontPath).isFile(), "探测到的字体文件应存在");
        }
    }

    /**
     * 生成一个包含中文段落的 docx 文件。
     *
     * @param file 目标 docx 文件
     * @throws IOException 写入失败时抛出
     */
    private static void writeChineseDocx(final File file) throws IOException {
        try (XWPFDocument document = new XWPFDocument();
             FileOutputStream out = new FileOutputStream(file)) {
            // xdocreport 要求文档含样式部件，POI 新建的空白文档默认没有，需显式创建（Word 正常保存的文档均自带）
            document.createStyles();
            final XWPFParagraph paragraph = document.createParagraph();
            paragraph.createRun().setText("中文转换测试：这是第一段中文，Hello PDF。");
            final XWPFParagraph second = document.createParagraph();
            second.createRun().setText("第二段：标点符号，。；！？以及数字 12345。");
            // 同理补上节属性（页面设置），xdocreport 依赖它读取页面尺寸；须置于正文末尾
            final CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();
            final CTPageSz pageSize = sectPr.addNewPgSz();
            pageSize.setW(BigInteger.valueOf(11906));
            pageSize.setH(BigInteger.valueOf(16838));
            document.write(out);
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
