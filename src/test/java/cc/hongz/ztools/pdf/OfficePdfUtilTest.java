package cc.hongz.ztools.pdf;

import cc.hongz.ztools.core.exceptions.UtilException;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.jupiter.api.Assumptions;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

@DisplayName("Office转PDF工具类测试")
public class OfficePdfUtilTest {

    /**
     * 测试用临时目录。
     */
    @TempDir
    private File tempDir;

    @Test
    @DisplayName("soffice探测结果合法测试")
    public void findSofficeTest() {
        final String soffice = OfficePdfUtil.findSoffice();

        if (soffice != null) {
            assertTrue(new File(soffice).isFile(), "探测到的 soffice 应存在");
        }
    }

    @Test
    @DisplayName("docx转PDF并落在目标目录测试")
    public void convertTest() throws IOException {
        Assumptions.assumeTrue(OfficePdfUtil.isAvailable(), "当前环境未安装 LibreOffice，跳过");

        final File docx = new File(this.tempDir, "office-source.docx");
        writeChineseDocx(docx);

        final File pdf = OfficePdfUtil.convert(docx, this.tempDir);

        assertTrue(pdf.isFile(), "PDF 文件应已生成");
        assertTrue(pdf.length() > 0, "PDF 文件不应为空");
        assertEquals("office-source.pdf", pdf.getName(), "产物应为同名 PDF");
    }

    @Test
    @DisplayName("源文件不存在时抛出异常测试")
    public void sourceNotExistsTest() {
        final File notExist = new File(this.tempDir, "not-exist.docx");

        assertThrows(UtilException.class, () -> OfficePdfUtil.convert(notExist, this.tempDir));
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
            document.createParagraph().createRun().setText("中文转换测试：LibreOffice 命令行方案。");
            document.write(out);
        }
    }
}
