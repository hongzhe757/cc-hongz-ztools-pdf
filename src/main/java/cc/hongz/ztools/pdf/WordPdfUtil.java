package cc.hongz.ztools.pdf;

import cc.hongz.ztools.core.exceptions.UtilException;
import cc.hongz.ztools.core.io.IORuntimeException;
import cc.hongz.ztools.core.util.StrUtil;
import com.lowagie.text.DocumentException;
import com.lowagie.text.Font;
import com.lowagie.text.pdf.BaseFont;
import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import fr.opensagres.xdocreport.openpdf.extension.font.IFontProvider;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.awt.Color;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

/**
 * Word（docx）转 PDF 工具类，基于 POI 与 xdocreport（OpenPDF 后端）实现纯 Java 转换。
 *
 * <p>与 {@link OfficePdfUtil} 的取舍：本类不依赖本机安装 LibreOffice，可嵌入任意环境，
 * 适合纯文本与常规图文混排的 docx；复杂排版（分栏、页眉页脚、浮动对象等）的还原度
 * 不及 LibreOffice 命令行方案，此类文档建议改用 {@link OfficePdfUtil}。</p>
 *
 * <p>中文字体：默认按操作系统探测常见中文字体文件（Windows 黑体/微软雅黑、Linux 文泉驿/Noto、
 * macOS 华文黑体/宋体），探测不到时不设置字体提供者，交由 OpenPDF 的系统字体注册表处理。
 * 如需固定字体，使用带 {@code fontPath} 参数的重载方法显式指定。</p>
 *
 * @author hongz
 * @since 1.0.0
 */
public final class WordPdfUtil {

    /**
     * 私有构造器，禁止实例化。
     */
    private WordPdfUtil() {
    }

    /**
     * 将 docx 文档转换为 PDF 文件，中文字体按操作系统自动探测。
     *
     * @param docx 源 docx 文件
     * @param pdf  目标 PDF 文件（不存在时自动创建，存在时覆盖）
     */
    public static void wordToPdf(final File docx, final File pdf) {
        wordToPdf(docx, pdf, detectChineseFont());
    }

    /**
     * 将 docx 文档转换为 PDF 文件。
     *
     * @param docx     源 docx 文件
     * @param pdf      目标 PDF 文件（不存在时自动创建，存在时覆盖）
     * @param fontPath 中文字体文件路径，为空时不设置字体提供者
     */
    public static void wordToPdf(final File docx, final File pdf, final String fontPath) {
        try (InputStream docxStream = new FileInputStream(docx);
             OutputStream pdfStream = new FileOutputStream(pdf)) {
            wordToPdf(docxStream, pdfStream, fontPath);
        } catch (IOException e) {
            throw new IORuntimeException(e);
        }
    }

    /**
     * 将 docx 文档流转换为 PDF 并写入输出流。
     *
     * @param docxStream 源 docx 输入流
     * @param pdfStream  目标 PDF 输出流
     * @param fontPath   中文字体文件路径，为空时不设置字体提供者
     */
    public static void wordToPdf(final InputStream docxStream, final OutputStream pdfStream, final String fontPath) {
        try (XWPFDocument document = new XWPFDocument(docxStream)) {
            final PdfOptions options = PdfOptions.create();
            if (StrUtil.isBlank(fontPath) == false) {
                // 字体只解析一次并在整个转换过程中复用，避免逐段重复读取字体文件
                options.fontProvider(new BaseFontProvider(createBaseFont(fontPath)));
            }
            PdfConverter.getInstance().convert(document, pdfStream, options);
        } catch (IOException e) {
            throw new IORuntimeException(e);
        }
    }

    /**
     * 探测当前操作系统的常见中文字体文件，详见 {@link PdfFontUtil#detectChineseFont()}。
     *
     * @return 探测到的字体文件路径，未探测到返回 {@code null}
     */
    public static String detectChineseFont() {
        return PdfFontUtil.detectChineseFont();
    }

    /**
     * 加载字体文件为 iText 基础字体。
     *
     * @param fontPath 字体文件路径
     * @return iText 基础字体
     * @throws UtilException 字体文件不存在或格式不支持时抛出
     */
    private static BaseFont createBaseFont(final String fontPath) {
        try {
            return BaseFont.createFont(fontPath, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
        } catch (DocumentException | IOException e) {
            throw new UtilException(e, "加载字体文件失败：{}", fontPath);
        }
    }

    /**
     * 复用同一个 {@link BaseFont} 实例的字体提供者，避免每次取字体都重新解析字体文件。
     */
    private static class BaseFontProvider implements IFontProvider {

        /**
         * 转换过程中复用的基础字体。
         */
        private final BaseFont baseFont;

        /**
         * 构造器。
         *
         * @param baseFont 转换过程中复用的基础字体
         */
        BaseFontProvider(final BaseFont baseFont) {
            this.baseFont = baseFont;
        }

        @Override
        public Font getFont(final String familyName, final String encoding, final float size, final int style,
                            final Color color) {
            final Font font = new Font(this.baseFont, size, style, color);
            if (familyName != null) {
                font.setFamily(familyName);
            }
            return font;
        }
    }
}
