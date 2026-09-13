package cc.hongz.ztools.pdf;

import cc.hongz.ztools.core.io.IORuntimeException;
import com.itextpdf.html2pdf.ConverterProperties;
import com.itextpdf.html2pdf.HtmlConverter;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

/**
 * PDF 工具类，基于 iText html2pdf 提供 HTML 转 PDF 能力。
 */
public class PdfUtil {

    /**
     * HTML 字符串转 PDF，写入输出流。
     *
     * @param html HTML 内容
     * @param out  输出流
     */
    public static void htmlToPdf(final String html, final OutputStream out) {
        htmlToPdf(html, out, null);
    }

    /**
     * HTML 字符串转 PDF，写入输出流。
     *
     * @param html  HTML 内容
     * @param out   输出流
     * @param props 转换属性（可设置 baseUri 等，可空）
     */
    public static void htmlToPdf(final String html, final OutputStream out, final ConverterProperties props) {
        if (props == null) {
            HtmlConverter.convertToPdf(html, out);
        } else {
            HtmlConverter.convertToPdf(html, out, props);
        }
    }

    /**
     * HTML 字符串转 PDF，写入文件。
     *
     * @param html HTML 内容
     * @param dest 目标 PDF 文件
     */
    public static void htmlToPdf(final String html, final File dest) {
        try (OutputStream out = new FileOutputStream(dest)) {
            htmlToPdf(html, out);
        } catch (IOException e) {
            throw new IORuntimeException(e);
        }
    }

    /**
     * HTML 文件转 PDF 文件。
     *
     * @param htmlFile HTML 源文件
     * @param dest     目标 PDF 文件
     */
    public static void htmlToPdf(final File htmlFile, final File dest) {
        try {
            HtmlConverter.convertToPdf(htmlFile, dest);
        } catch (IOException e) {
            throw new IORuntimeException(e);
        }
    }

    /**
     * HTML 字符串转 PDF 字节数组。
     *
     * @param html HTML 内容
     * @return PDF 字节数组
     */
    public static byte[] htmlToPdfBytes(final String html) {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        htmlToPdf(html, out);
        return out.toByteArray();
    }
}
