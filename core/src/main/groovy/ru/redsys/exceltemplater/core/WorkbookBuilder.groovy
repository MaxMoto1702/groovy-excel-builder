package ru.redsys.exceltemplater.core

import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.streaming.SXSSFWorkbook
import org.slf4j.LoggerFactory

class WorkbookBuilder {
    final static LOGGER = LoggerFactory.getLogger(WorkbookBuilder)
    final static STYLE_ATTRIBUTES = ['border', 'backgroundColor', 'align', 'verticalAlign']

    Workbook workbook
    Styler styler
    private def currentSheet
    private def currentRow
    private def currentCell
    private def currentRowIndex
    private def currentCellIndex
    Map<List, CellStyle> mixes

    WorkbookBuilder(workbook = null) {
        this.workbook = workbook
        this.mixes = new HashMap<>()
    }

    def workbook(Closure closure) {
        closure.delegate = this
        LOGGER.debug "Create Excel book (Stream workbook)"
        this.workbook = this.workbook ?: new SXSSFWorkbook()
        closure()
        this
    }

    def sheet(String name = null, Closure closure) {
        closure.delegate = this
        if (!name) name = generateSheetName()
        LOGGER.debug "Create sheet with name '$name'"
        currentSheet = workbook.getSheet(name) ?: workbook.createSheet(name)
        currentRowIndex = -1
        closure()
        currentSheet = null
        this
    }

    def row(Closure closure) {
        row(null, closure)
    }

    def row(Map attributes = null, Closure closure = null) {
        currentRowIndex++
        currentRow = currentSheet.createRow currentRowIndex
        if (attributes?.height) currentRow.heightInPoints = attributes.height
        currentCellIndex = -1
        if (closure) {
            closure.delegate = this
            closure()
        }
        this
    }

    def cell(Closure closure) {
        cell(null, closure)
    }

    def cell(Map attributes = null, Closure closure = null) {
        currentCellIndex++
        closure.delegate = this
        currentCell = currentRow.createCell currentCellIndex
        if (attributes?.rowspan || attributes?.colspan) createMergeRegion(attributes.rowspan, attributes.colspan)
        def styleAttributes = attributes.findAll { attribute -> attribute.key in STYLE_ATTRIBUTES }
        if (styleAttributes) applyStyle(styleAttributes.values())
        def v = closure()
        currentCell.setCellValue "$v"
        this
    }

    private def applyStyle(Collection collection) {
        currentCell.setCellStyle(mix(collection))
    }

    private CellStyle mix(Collection styleMix) {
        mixes.containsKey(styleMix) ? mixes.get(styleMix) : createMix(styleMix)
    }

    private def createMix(Collection styleMix) {
        CellStyle newStyle = workbook.createCellStyle();
        for (Short styleKey : styleMix) {
            CellStyle tempStyle = mixes.get(Arrays.asList(styleKey));
            mixStyles(newStyle, tempStyle);
        }
        mixes.put(styleMix, newStyle);
        return newStyle;
    }

     static final int DEFAULT_ALIGNMENT_VALUE = 0
     static final int DEFAULT_BORDER_VALUE = 0
     static final int DEFAULT_BORDER_COLOR_VALUE = 8
     static final int DEFAULT_DATA_FORMAT_VALUE = 0
     static final int DEFAULT_GROUND_COLOR_VALUE = 64
     static final int DEFAULT_FILL_PATTERN_VALUE = 0
     static final int DEFAULT_FONT_VALUE = 0
     static final int DEFAULT_VERTICAL_ALIGNMENT_VALUE = 2
     static final int DEFAULT_INDENTION_VALUE = 0
     static final int DEFAULT_ROTATION_VALUE = 0

    private def mixStyles(CellStyle mainStyle, CellStyle tempStyle) {
        if (tempStyle.getAlignment() != DEFAULT_ALIGNMENT_VALUE) {
            mainStyle.setAlignment(tempStyle.getAlignment());
        }
        if (tempStyle.getBorderTop() != DEFAULT_BORDER_VALUE) {
            mainStyle.setBorderTop(tempStyle.getBorderTop());
        }
        if (tempStyle.getBorderLeft() != DEFAULT_BORDER_VALUE) {
            mainStyle.setBorderLeft(tempStyle.getBorderLeft());
        }
        if (tempStyle.getBorderRight() != DEFAULT_BORDER_VALUE) {
            mainStyle.setBorderRight(tempStyle.getBorderRight());
        }
        if (tempStyle.getBorderBottom() != DEFAULT_BORDER_VALUE) {
            mainStyle.setBorderBottom(tempStyle.getBorderBottom());
        }
        if (tempStyle.getTopBorderColor() != DEFAULT_BORDER_COLOR_VALUE) {
            mainStyle.setTopBorderColor(tempStyle.getTopBorderColor());
        }
        if (tempStyle.getLeftBorderColor() != DEFAULT_BORDER_COLOR_VALUE) {
            mainStyle.setLeftBorderColor(tempStyle.getLeftBorderColor());
        }
        if (tempStyle.getRightBorderColor() != DEFAULT_BORDER_COLOR_VALUE) {
            mainStyle.setRightBorderColor(tempStyle.getRightBorderColor());
        }
        if (tempStyle.getBottomBorderColor() != DEFAULT_BORDER_COLOR_VALUE) {
            mainStyle.setBottomBorderColor(tempStyle.getBottomBorderColor());
        }
        if (tempStyle.getDataFormat() != DEFAULT_DATA_FORMAT_VALUE) {
            mainStyle.setDataFormat(tempStyle.getDataFormat());
        }
        if (tempStyle.getFillBackgroundColor() != DEFAULT_GROUND_COLOR_VALUE) {
            mainStyle.setFillBackgroundColor(tempStyle.getFillBackgroundColor());
        }
        if (tempStyle.getFillForegroundColor() != DEFAULT_GROUND_COLOR_VALUE) {
            mainStyle.setFillForegroundColor(tempStyle.getFillBackgroundColor());
        }
        if (tempStyle.getFillPattern() != DEFAULT_FILL_PATTERN_VALUE) {
            mainStyle.setFillPattern(tempStyle.getFillPattern());
        }
        if (tempStyle.getFontIndex() != DEFAULT_FONT_VALUE) {
            mainStyle.setFont(workbook.getFontAt(tempStyle.getFontIndex()));
        }
        if (tempStyle.getIndention() != DEFAULT_INDENTION_VALUE) {
            mainStyle.setIndention(tempStyle.getIndention());
        }
        if (tempStyle.getRotation() != DEFAULT_ROTATION_VALUE) {
            mainStyle.setRotation(tempStyle.getRotation());
        }
        if (tempStyle.getVerticalAlignment() != DEFAULT_VERTICAL_ALIGNMENT_VALUE) {
            mainStyle.setVerticalAlignment(tempStyle.getVerticalAlignment());
        }
        if (tempStyle.getWrapText()) {
            mainStyle.setWrapText(true);
        }
    }

    def styles(Closure closure) {
        closure.delegate = this
        closure()
        this
    }

    def style(Map attributes, Closure closure) {
        closure()
        this
    }

    private def createMergeRegion(rowspan, colspan) {
        currentSheet.addMergedRegion(new CellRangeAddress(currentRowIndex, currentRowIndex + (rowspan ?: 1) - 1, currentCellIndex, currentCellIndex + (colspan ?: 1) - 1))
    }

    private String generateSheetName() { "Sheet $workbook.numberOfSheets" }
}
