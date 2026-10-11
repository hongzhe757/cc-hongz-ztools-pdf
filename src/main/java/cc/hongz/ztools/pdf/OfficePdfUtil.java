package cc.hongz.ztools.pdf;

import cc.hongz.ztools.core.exceptions.UtilException;
import cc.hongz.ztools.core.util.RuntimeUtil;

import java.io.File;
import java.util.concurrent.TimeUnit;

/**
 * Office 文档转 PDF 工具类，调用本机 LibreOffice（soffice）命令行完成转换。
 *
 * <p>支持 docx、xlsx、pptx 等 LibreOffice 可识别的格式，复杂排版的还原度高于纯 Java 方案，
 * 但要求运行环境已安装 LibreOffice。未安装时可用 {@link #isAvailable()} 判断并降级到
 * {@link WordPdfUtil}（纯 Java，无需外部程序）。</p>
 *
 * <p>并发安全：每次调用都会分配独立的 LibreOffice 用户配置目录，避免多实例共用配置目录
 * 导致的单实例锁冲突。</p>
 *
 * @author hongz
 * @since 1.0.0
 */
public final class OfficePdfUtil {

    /**
     * soffice 可执行文件的常见安装路径，按优先级排列。
     */
    private static final String[] SOFFICE_CANDIDATES = {
            "C:/Program Files/LibreOffice/program/soffice.exe",
            "C:/Program Files (x86)/LibreOffice/program/soffice.exe",
            "/usr/bin/soffice",
            "/usr/local/bin/soffice",
            "/opt/libreoffice/program/soffice",
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    };

    /**
     * 在系统 PATH 中查找 soffice 时使用的可执行文件名。
     */
    private static final String[] SOFFICE_COMMAND_NAMES = {"soffice", "soffice.exe"};

    /**
     * 默认转换超时时间。
     */
    private static final long DEFAULT_TIMEOUT_MILLIS = TimeUnit.MINUTES.toMillis(2);

    /**
     * 私有构造器，禁止实例化。
     */
    private OfficePdfUtil() {
    }

    /**
     * 将 Office 文档转换为 PDF，使用默认超时与自动探测的 soffice。
     *
     * @param src     源文档
     * @param destDir 目标目录（不存在时自动创建）
     * @return 生成的 PDF 文件
     * @throws UtilException 未找到 soffice、转换超时或转换失败时抛出
     */
    public static File convert(final File src, final File destDir) {
        return convert(src, destDir, findSoffice(), DEFAULT_TIMEOUT_MILLIS);
    }

    /**
     * 将 Office 文档转换为 PDF，使用指定 soffice 与默认超时。
     *
     * @param src         源文档
     * @param destDir     目标目录（不存在时自动创建）
     * @param sofficePath soffice 可执行文件路径，为 {@code null} 时自动探测
     * @return 生成的 PDF 文件
     * @throws UtilException 未找到 soffice、转换超时或转换失败时抛出
     */
    public static File convert(final File src, final File destDir, final String sofficePath) {
        return convert(src, destDir, sofficePath, DEFAULT_TIMEOUT_MILLIS);
    }

    /**
     * 将 Office 文档转换为 PDF。
     *
     * @param src         源文档
     * @param destDir     目标目录（不存在时自动创建）
     * @param sofficePath soffice 可执行文件路径，为 {@code null} 时自动探测
     * @param timeoutMs   转换超时时间（毫秒）
     * @return 生成的 PDF 文件
     * @throws UtilException 未找到 soffice、转换超时或转换失败时抛出
     */
    public static File convert(final File src, final File destDir, final String sofficePath, final long timeoutMs) {
        if (src == null || src.isFile() == false) {
            throw new UtilException("源文件不存在：{}", src);
        }
        if (destDir == null) {
            throw new UtilException("目标目录不能为空。");
        }
        if (destDir.isDirectory() == false && (destDir.exists() || destDir.mkdirs() == false)) {
            throw new UtilException("目标目录不可用：{}", destDir);
        }
        final String executable = sofficePath == null ? findSoffice() : sofficePath;
        if (executable == null) {
            throw new UtilException("未找到 LibreOffice（soffice），请先安装或通过 sofficePath 参数指定。");
        }

        final File target = new File(destDir, baseName(src.getName()) + ".pdf");
        // 先清理旧产物，避免转换失败时被历史文件误判为成功
        if (target.exists() && target.delete() == false) {
            throw new UtilException("删除已存在的目标文件失败：{}", target);
        }

        final String[] commands = {
                executable,
                "--headless",
                "--norestore",
                "--convert-to", "pdf",
                "--outdir", destDir.getAbsolutePath(),
                "-env:UserInstallation=" + createProfileDir().toURI(),
                src.getAbsolutePath(),
        };
        run(commands, src, timeoutMs);

        if (target.isFile() && target.length() > 0) {
            return target;
        }
        throw new UtilException("转换失败，未生成 PDF 文件：{}", target);
    }

    /**
     * 判断当前环境是否可用 LibreOffice 转换。
     *
     * @return 能找到 soffice 可执行文件返回 {@code true}
     */
    public static boolean isAvailable() {
        return findSoffice() != null;
    }

    /**
     * 探测 soffice 可执行文件：先查常见安装路径，再查系统 PATH。
     *
     * @return soffice 可执行文件路径，未找到返回 {@code null}
     */
    public static String findSoffice() {
        for (String candidate : SOFFICE_CANDIDATES) {
            if (new File(candidate).isFile()) {
                return candidate;
            }
        }
        for (String commandName : SOFFICE_COMMAND_NAMES) {
            final File found = findOnPath(commandName);
            if (found != null) {
                return found.getAbsolutePath();
            }
        }
        return null;
    }

    /**
     * 执行转换命令并等待其结束。
     *
     * @param commands  命令及参数
     * @param src       源文档，仅用于异常信息
     * @param timeoutMs 超时时间（毫秒）
     */
    private static void run(final String[] commands, final File src, final long timeoutMs) {
        final Process process = RuntimeUtil.exec(commands);
        final boolean finished;
        try {
            finished = process.waitFor(timeoutMs, TimeUnit.MILLISECONDS);
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
            RuntimeUtil.destroy(process);
            throw new UtilException(e, "转换被中断：{}", src);
        }
        if (finished == false) {
            RuntimeUtil.destroy(process);
            throw new UtilException("转换超时（{} 毫秒）：{}", timeoutMs, src);
        }
    }

    /**
     * 创建本次转换独占的 LibreOffice 用户配置目录。
     *
     * @return 配置目录
     */
    private static File createProfileDir() {
        final File profileDir = new File(System.getProperty("java.io.tmpdir"),
                "ztools-libreoffice-" + System.nanoTime());
        if (profileDir.mkdirs() == false && profileDir.isDirectory() == false) {
            throw new UtilException("创建 LibreOffice 配置目录失败：{}", profileDir);
        }
        return profileDir;
    }

    /**
     * 在系统 PATH 中查找可执行文件。
     *
     * @param commandName 可执行文件名
     * @return 找到的文件，未找到返回 {@code null}
     */
    private static File findOnPath(final String commandName) {
        final String pathEnv = System.getenv("PATH");
        if (pathEnv == null || pathEnv.isEmpty()) {
            return null;
        }
        for (String dir : pathEnv.split(File.pathSeparator)) {
            if (dir.isEmpty()) {
                continue;
            }
            final File candidate = new File(dir, commandName);
            if (candidate.isFile()) {
                return candidate;
            }
        }
        return null;
    }

    /**
     * 取文件名中最后一个点之前的部分，无扩展名时返回原名。
     *
     * @param fileName 文件名
     * @return 不含扩展名的文件名
     */
    private static String baseName(final String fileName) {
        final int dotIndex = fileName.lastIndexOf('.');
        return dotIndex > 0 ? fileName.substring(0, dotIndex) : fileName;
    }
}
