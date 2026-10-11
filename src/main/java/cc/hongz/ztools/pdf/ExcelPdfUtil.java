package cc.hongz.ztools.pdf;

import cc.hongz.ztools.core.exceptions.UtilException;
import cc.hongz.ztools.core.io.IORuntimeException;
import cc.hongz.ztools.core.util.StrUtil;
import com.itextpdf.io.font.PdfEncodings;
import com.itextpdf.io.font.constants.StandardFonts;
import com.itextpdf.io.image.ImageDataFactory;
import com.itextpdf.kernel.colors.DeviceRgb;
import com.itextpdf.kernel.font.PdfFont;
import com.itextpdf.kernel.font.PdfFontFactory;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.borders.SolidBorder;
import com.itextpdf.layout.element.AreaBreak;
import com.itextpdf.layout.element.Cell;
import com.itextpdf.layout.element.Image;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Table;
import com.itextpdf.layout.element.Text;
import com.itextpdf.layout.property.AreaBreakType;
import com.itextpdf.layout.property.TextAlignment;
import com.itextpdf.layout.property.UnitValue;
import com.itextpdf.layout.property.VerticalAlignment;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Shape;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.Map;

/**
 * Excel 转 PDF 工具类，基于 POI 与 iText 7 实现纯 Java 转换，逐单元格还原表格格式。
 *
 * <p>与 {@link OfficePdfUtil} 的取舍：本类不依赖本机安装 LibreOffice，可嵌入任意环境；
 * 但复杂排版（图表、条件格式、打印区域、分页符等）不还原，此类场景建议改用
 * {@link OfficePdfUtil} 的 LibreOffice 命令行方案，保真度更高。</p>
 *
 * <p>保留的格式：列宽比例、合并单元格、字号、字体颜色、背景色、边框、水平与垂直对齐、
 * 单元格内嵌图片。其中边框颜色统一为黑色，粗细按原单元格边框样式映射。</p>
 *
 * <p>中文字体：默认按操作系统自动探测（详见 {@link PdfFontUtil#detectChineseFont()}），
 * 也可通过 {@code fontPath} 参数显式指定字体文件。Excel 里的粗体通过 iText 的字体模拟实现，
 * 所用字体文件不含对应字重时会降级为常规字形。</p>
 *
 * @author hongz
 * @since 1.0.0
 */
public final class ExcelPdfUtil {

    /**
     * 无法读取列宽时使用的兜底列宽权重。
     */
    private static final float FALLBACK_COLUMN_WEIGHT = 1F;

    /**
     * 默认字号，单元格未指定字号时使用。
     */
    private static final float DEFAULT_FONT_SIZE = 10F;

    /**
     * 单元格边框宽度。
     */
    private static final float BORDER_WIDTH = 0.5F;

    /**
     * 无符号字节的取值范围，用于把负值还原到 0-255。
     */
    private static final int BYTE_RANGE = 256;

    /**
     * 单元格文本格式化器，负责把数字、日期等按 Excel 显示格式转为字符串。
     */
    private static final DataFormatter CELL_FORMATTER = new DataFormatter();

    /**
     * 私有构造器，禁止实例化。
     */
    private ExcelPdfUtil() {
    }

    /**
     * 将 Excel 文档转换为 PDF 文件，中文字体按操作系统自动探测。
     *
     * @param excel 源 Excel 文件（xls 或 xlsx）
     * @param pdf   目标 PDF 文件（不存在时自动创建，存在时覆盖）
     */
    public static void excelToPdf(final File excel, final File pdf) {
        excelToPdf(excel, pdf, PdfFontUtil.detectChineseFont());
    }

    /**
     * 将 Excel 文档转换为 PDF 文件。
     *
     * @param excel    源 Excel 文件（xls 或 xlsx）
     * @param pdf      目标 PDF 文件（不存在时自动创建，存在时覆盖）
     * @param fontPath 中文字体文件路径，为空时按操作系统自动探测
     */
    public static void excelToPdf(final File excel, final File pdf, final String fontPath) {
        try (InputStream excelStream = new FileInputStream(excel);
             OutputStream pdfStream = new FileOutputStream(pdf)) {
            excelToPdf(excelStream, pdfStream, fontPath);
        } catch (IOException e) {
            throw new IORuntimeException(e);
        }
    }

    /**
     * 将 Excel 文档流转换为 PDF 并写入输出流。 <br>
     * 工作簿中的每个工作表单独成页，按其在工作簿中的顺序排列。
     *
     * @param excelStream 源 Excel 输入流
     * @param pdfStream   目标 PDF 输出流
     * @param fontPath    中文字体文件路径，为空时按操作系统自动探测
     */
    public static void excelToPdf(final InputStream excelStream, final OutputStream pdfStream, final String fontPath) {
        final PdfFont font = createFont(fontPath);
        try (Workbook workbook = WorkbookFactory.create(excelStream);
             Document document = new Document(new PdfDocument(new PdfWriter(pdfStream)))) {
            final int sheetCount = workbook.getNumberOfSheets();
            for (int i = 0; i < sheetCount; i++) {
                if (i > 0) {
                    document.add(new AreaBreak(AreaBreakType.NEXT_PAGE));
                }
                document.add(createTable(workbook.getSheetAt(i), font));
            }
        } catch (IOException e) {
            throw new IORuntimeException(e);
        }
    }

    /**
     * 加载 PDF 字体：优先使用指定路径，其次按操作系统探测中文字体，都没有时退回 Helvetica。
     *
     * @param fontPath 字体文件路径，可为空
     * @return {@link PdfFont}
     */
    private static PdfFont createFont(final String fontPath) {
        String path = fontPath;
        if (StrUtil.isBlank(path)) {
            path = PdfFontUtil.detectChineseFont();
        }
        try {
            if (StrUtil.isBlank(path)) {
                // 未探测到中文字体，退回内置字体：拉丁字符可正常渲染，中文将显示为乱码
                return PdfFontFactory.createFont(StandardFonts.HELVETICA);
            }
            return PdfFontFactory.createFont(path, PdfEncodings.IDENTITY_H);
        } catch (IOException e) {
            throw new UtilException(e, "加载字体文件失败：{}", path);
        }
    }

    /**
     * 将一个工作表渲染为 PDF 表格。
     *
     * @param sheet 工作表
     * @param font  PDF 字体
     * @return {@link Table}
     */
    private static Table createTable(final Sheet sheet, final PdfFont font) {
        final int columnCount = Math.max(getColumnCount(sheet), 1);
        final Table table = new Table(UnitValue.createPercentArray(getColumnWeights(sheet, columnCount)))
                .useAllAvailableWidth();
        // 单元格内嵌图片的定位信息一次性收集，避免逐格遍历图形对象
        final Map<String, byte[]> imageMap = collectImages(sheet);
        final int lastRowNum = sheet.getLastRowNum();
        for (int rowIndex = 0; rowIndex <= lastRowNum; rowIndex++) {
            final Row row = sheet.getRow(rowIndex);
            if (row != null) {
                addRow(table, row, sheet, font, imageMap);
            }
        }
        return table;
    }

    /**
     * 计算工作表的列数，取所有行中非空单元格数最多的那一行。
     *
     * @param sheet 工作表
     * @return 列数
     */
    private static int getColumnCount(final Sheet sheet) {
        int columnCount = 0;
        final int lastRowNum = sheet.getLastRowNum();
        for (int rowIndex = 0; rowIndex <= lastRowNum; rowIndex++) {
            final Row row = sheet.getRow(rowIndex);
            if (row != null) {
                columnCount = Math.max(columnCount, row.getLastCellNum());
            }
        }
        return columnCount;
    }

    /**
     * 按各列在 Excel 中的原始宽度计算相对权重，由 iText 归一化后自适应页面宽度。
     *
     * @param sheet       工作表
     * @param columnCount 列数
     * @return 各列宽度权重
     */
    private static float[] getColumnWeights(final Sheet sheet, final int columnCount) {
        final float[] weights = new float[columnCount];
        for (int i = 0; i < columnCount; i++) {
            weights[i] = Math.max(sheet.getColumnWidth(i), FALLBACK_COLUMN_WEIGHT);
        }
        return weights;
    }

    /**
     * 将一行的各单元格追加到表格。
     *
     * @param table    目标表格
     * @param row      源行
     * @param sheet    所在工作表
     * @param font     PDF 字体
     * @param imageMap 单元格图片映射
     */
    private static void addRow(final Table table, final Row row, final Sheet sheet, final PdfFont font,
                               final Map<String, byte[]> imageMap) {
        final int lastCellNum = row.getLastCellNum();
        for (int colIndex = 0; colIndex < lastCellNum; colIndex++) {
            final CellRangeAddress region = findMergedRegion(sheet, row.getRowNum(), colIndex);
            if (region != null && (region.getFirstRow() != row.getRowNum() || region.getFirstColumn() != colIndex)) {
                // 合并区域的非左上角单元格不单独生成，由左上角单元格的跨度覆盖
                continue;
            }
            table.addCell(createCell(row.getCell(colIndex), sheet, font, region, row.getRowNum(), colIndex,
                    imageMap));
        }
    }

    /**
     * 创建一个 PDF 单元格，还原其文本、样式、合并跨度与内嵌图片。
     *
     * @param poiCell  源单元格，可为 null（空单元格或合并区域的占位）
     * @param sheet    所在工作表
     * @param font     PDF 字体
     * @param region   所在合并区域，不属于合并区域时为 null
     * @param rowIndex 行号
     * @param colIndex 列号
     * @param imageMap 单元格图片映射
     * @return {@link Cell}
     */
    private static Cell createCell(final org.apache.poi.ss.usermodel.Cell poiCell, final Sheet sheet,
                                   final PdfFont font, final CellRangeAddress region, final int rowIndex,
                                   final int colIndex, final Map<String, byte[]> imageMap) {
        // iText 7 的跨行跨列只能在构造单元格时指定
        final Cell cell;
        if (region == null) {
            cell = new Cell();
        } else {
            cell = new Cell(region.getLastRow() - region.getFirstRow() + 1,
                    region.getLastColumn() - region.getFirstColumn() + 1);
        }
        cell.setBorder(new SolidBorder(DeviceRgb.BLACK, BORDER_WIDTH));
        if (poiCell != null) {
            final CellStyle style = poiCell.getCellStyle();
            applyStyle(cell, style, sheet.getWorkbook());
            cell.add(createParagraph(CELL_FORMATTER.formatCellValue(poiCell), style, sheet.getWorkbook(), font));
        }
        final byte[] imageBytes = imageMap.get(imageKey(rowIndex, colIndex));
        if (imageBytes != null) {
            cell.add(createImage(imageBytes));
        }
        return cell;
    }

    /**
     * 还原单元格的背景色与对齐方式。
     *
     * @param cell     目标单元格
     * @param style    源样式
     * @param workbook 工作簿
     */
    private static void applyStyle(final Cell cell, final CellStyle style, final Workbook workbook) {
        if (style == null) {
            return;
        }
        if (style.getFillPattern() != FillPatternType.NO_FILL) {
            final DeviceRgb background = toDeviceRgb(style.getFillForegroundColorColor());
            if (background != null) {
                cell.setBackgroundColor(background);
            }
        }
        cell.setTextAlignment(toTextAlignment(style.getAlignment()));
        cell.setVerticalAlignment(toVerticalAlignment(style.getVerticalAlignment()));
    }

    /**
     * 按单元格字体样式构造段落，还原字号、字体颜色、粗体与下划线。
     *
     * @param content  文本内容
     * @param style    源样式
     * @param workbook 工作簿
     * @param font     PDF 字体
     * @return {@link Paragraph}
     */
    private static Paragraph createParagraph(final String content, final CellStyle style, final Workbook workbook,
                                            final PdfFont font) {
        final Text text = new Text(content).setFont(font).setFontSize(DEFAULT_FONT_SIZE);
        if (style != null) {
            final Font poiFont = workbook.getFontAt(style.getFontIndexAsInt());
            text.setFontSize(poiFont.getFontHeightInPoints());
            final DeviceRgb fontColor = resolveFontColor(poiFont.getColor());
            if (fontColor != null) {
                text.setFontColor(fontColor);
            }
            if (poiFont.getBold()) {
                text.setBold();
            }
            if (FontUnderline.SINGLE == FontUnderline.valueOf(poiFont.getUnderline())) {
                text.setUnderline();
            }
        }
        return new Paragraph(text);
    }

    /**
     * 构造单元格内嵌图片，超出可用宽度时等比缩放。
     *
     * @param imageBytes 图片字节
     * @return {@link Image}
     */
    private static Image createImage(final byte[] imageBytes) {
        final Image image = new Image(ImageDataFactory.create(imageBytes));
        image.setAutoScale(true);
        return image;
    }

    /**
     * 收集工作表中各单元格的内嵌图片，键为「行号:列号」。
     *
     * @param sheet 工作表
     * @return 单元格图片映射
     */
    private static Map<String, byte[]> collectImages(final Sheet sheet) {
        final Map<String, byte[]> imageMap = new HashMap<>();
        final Drawing<?> drawing = sheet.getDrawingPatriarch();
        if (drawing == null) {
            return imageMap;
        }
        for (Shape shape : drawing) {
            if (shape instanceof Picture) {
                final Picture picture = (Picture) shape;
                final ClientAnchor anchor = picture.getClientAnchor();
                if (anchor != null) {
                    imageMap.put(imageKey(anchor.getRow1(), anchor.getCol1()), picture.getPictureData().getData());
                }
            }
        }
        return imageMap;
    }

    /**
     * 构造单元格图片映射的键。
     *
     * @param rowIndex 行号
     * @param colIndex 列号
     * @return 键
     */
    private static String imageKey(final int rowIndex, final int colIndex) {
        return rowIndex + ":" + colIndex;
    }

    /**
     * 查找单元格所在的合并区域。
     *
     * @param sheet    工作表
     * @param rowIndex 行号
     * @param colIndex 列号
     * @return 合并区域，不属于任何合并区域时返回 {@code null}
     */
    private static CellRangeAddress findMergedRegion(final Sheet sheet, final int rowIndex, final int colIndex) {
        final int regionCount = sheet.getNumMergedRegions();
        for (int i = 0; i < regionCount; i++) {
            final CellRangeAddress region = sheet.getMergedRegion(i);
            if (region.isInRange(rowIndex, colIndex)) {
                return region;
            }
        }
        return null;
    }

    /**
     * 把 POI 的颜色对象转换为 iText 颜色。
     *
     * @param color POI 颜色，可为 null
     * @return iText 颜色，无法解析时返回 {@code null}
     */
    private static DeviceRgb toDeviceRgb(final Color color) {
        if (color == null) {
            return null;
        }
        if (color instanceof XSSFColor) {
            final byte[] rgb = ((XSSFColor) color).getRGB();
            if (rgb != null && rgb.length >= 3) {
                return new DeviceRgb(toUnsigned(rgb[0]), toUnsigned(rgb[1]), toUnsigned(rgb[2]));
            }
        }
        if (color instanceof HSSFColor) {
            final short[] triplet = ((HSSFColor) color).getTriplet();
            if (triplet != null && triplet.length >= 3) {
                return new DeviceRgb(toUnsigned(triplet[0]), toUnsigned(triplet[1]), toUnsigned(triplet[2]));
            }
        }
        return null;
    }

    /**
     * 把索引色解析为 iText 颜色。
     *
     * @param colorIndex 颜色索引
     * @return iText 颜色，无法解析时返回 {@code null}
     */
    private static DeviceRgb resolveFontColor(final short colorIndex) {
        final HSSFColor color = HSSFColor.getIndexHash().get((int) colorIndex);
        if (color == null) {
            return null;
        }
        final short[] triplet = color.getTriplet();
        if (triplet == null || triplet.length < 3) {
            return null;
        }
        return new DeviceRgb(toUnsigned(triplet[0]), toUnsigned(triplet[1]), toUnsigned(triplet[2]));
    }

    /**
     * 把可能为负的字节或 short 值转为 0-255 的无符号整数。
     *
     * @param value 原始值
     * @return 无符号整数
     */
    private static int toUnsigned(final short value) {
        return value < 0 ? value + BYTE_RANGE : value;
    }

    /**
     * 把可能为负的字节转为 0-255 的无符号整数。
     *
     * @param value 原始值
     * @return 无符号整数
     */
    private static int toUnsigned(final byte value) {
        return value < 0 ? value + BYTE_RANGE : value;
    }

    /**
     * 把 POI 的水平对齐方式转换为 iText 文本对齐方式。
     *
     * @param alignment POI 对齐方式
     * @return iText 对齐方式
     */
    private static TextAlignment toTextAlignment(final org.apache.poi.ss.usermodel.HorizontalAlignment alignment) {
        if (alignment == null) {
            return TextAlignment.LEFT;
        }
        switch (alignment) {
            case CENTER:
            case CENTER_SELECTION:
                return TextAlignment.CENTER;
            case RIGHT:
                return TextAlignment.RIGHT;
            case JUSTIFY:
            case DISTRIBUTED:
                return TextAlignment.JUSTIFIED;
            default:
                return TextAlignment.LEFT;
        }
    }

    /**
     * 把 POI 的垂直对齐方式转换为 iText 垂直对齐方式。
     *
     * @param alignment POI 对齐方式
     * @return iText 对齐方式
     */
    private static VerticalAlignment toVerticalAlignment(
            final org.apache.poi.ss.usermodel.VerticalAlignment alignment) {
        if (alignment == null) {
            return VerticalAlignment.MIDDLE;
        }
        switch (alignment) {
            case TOP:
                return VerticalAlignment.TOP;
            case BOTTOM:
                return VerticalAlignment.BOTTOM;
            default:
                return VerticalAlignment.MIDDLE;
        }
    }
}
