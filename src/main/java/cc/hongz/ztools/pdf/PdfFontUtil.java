package cc.hongz.ztools.pdf;

import java.io.File;

/**
 * PDF 中文字体探测工具类，按操作系统查找常见的中文字体文件。
 *
 * <p>本类只依赖 JDK，不引入任何 PDF / Office 三方库，便于被 {@link PdfUtil}、
 * {@link WordPdfUtil}、{@link ExcelPdfUtil} 等各类转换能力共用，避免因共用字体探测
 * 而互相拖入对方的三方依赖。</p>
 *
 * @author hongz
 * @since 1.0.0
 */
public final class PdfFontUtil {

    /**
     * Windows 常见中文字体文件，按优先级排列。
     */
    private static final String[] WINDOWS_FONT_CANDIDATES = {
            "C:/Windows/Fonts/simhei.ttf",
            "C:/Windows/Fonts/msyh.ttc",
            "C:/Windows/Fonts/simsun.ttc",
    };

    /**
     * Linux 常见中文字体文件，按优先级排列。
     */
    private static final String[] LINUX_FONT_CANDIDATES = {
            "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
            "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
            "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
            "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
            "/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc",
    };

    /**
     * macOS 常见中文字体文件，按优先级排列。
     */
    private static final String[] MAC_FONT_CANDIDATES = {
            "/System/Library/Fonts/STHeiti Light.ttc",
            "/System/Library/Fonts/Songti.ttc",
            "/System/Library/Fonts/PingFang.ttc",
    };

    /**
     * 私有构造器，禁止实例化。
     */
    private PdfFontUtil() {
    }

    /**
     * 探测当前操作系统的常见中文字体文件。
     *
     * @return 探测到的字体文件路径，未探测到返回 {@code null}
     */
    public static String detectChineseFont() {
        final String osName = System.getProperty("os.name");
        final String[] candidates;
        if (osName != null && osName.toLowerCase().contains("win")) {
            candidates = WINDOWS_FONT_CANDIDATES;
        } else if (osName != null && osName.toLowerCase().contains("mac")) {
            candidates = MAC_FONT_CANDIDATES;
        } else {
            candidates = LINUX_FONT_CANDIDATES;
        }
        for (String candidate : candidates) {
            if (new File(candidate).isFile()) {
                return candidate;
            }
        }
        return null;
    }
}
