package net.serebryansky.excel_builder

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.streaming.SXSSFWorkbook
import org.slf4j.LoggerFactory
import java.util.Calendar
import java.util.Date

class ExcelBuilder {
    companion object {
        private val logger = LoggerFactory.getLogger(ExcelBuilder::class.java)

        @JvmStatic
        var defaultRowHeight: Float = -1f
    }

    var workbook: SXSSFWorkbook = SXSSFWorkbook()

    private val fonts = mutableMapOf<String, Font>()
    private val formats = mutableMapOf<String, Short>()
    private val styles = mutableMapOf<String, CellStyle>()

    private var currentSheet: Sheet? = null
    private var currentRowIndex = -1
    private var currentRow: Row? = null
    private var currentRowStyle: CellStyle? = null
    private var currentCellIndex = -1
    private var currentCell: Cell? = null
    private var currentCellStyle: CellStyle? = null

    fun config(block: ExcelBuilder.() -> Unit): ExcelBuilder {
        logger.debug("Start config builder")
        block()
        logger.debug("Finish config builder")
        return this
    }

    fun workbook(block: () -> SXSSFWorkbook): ExcelBuilder {
        workbook = block()
        return this
    }

    fun style(
        styleNameObj: Any?,
        block: ExcelBuilder.(CellStyle) -> Unit,
    ): CellStyle {
        val styleName = styleNameObj?.toString() ?: throw RuntimeException("Style name is null")
        logger.debug("Create style with name: {}", styleName)
        val newStyle = workbook.createCellStyle()
        block(newStyle)
        styles[styleName] = newStyle
        return newStyle
    }

    fun font(fontNameObj: Any?): Font {
        val fontName = fontNameObj?.toString() ?: throw RuntimeException("Font name is null")
        return fonts[fontName] ?: throw RuntimeException("Font $fontName not found")
    }

    fun font(
        fontNameObj: Any?,
        block: ExcelBuilder.(Font) -> Unit,
    ): Font {
        val fontName = fontNameObj?.toString() ?: throw RuntimeException("Font name is null")
        logger.debug("Create font with name: {}", fontName)
        val newFont = workbook.createFont()
        block(newFont)
        fonts[fontName] = newFont
        return newFont
    }

    fun dataFormat(formatNameObj: Any?): Short {
        val formatName = formatNameObj?.toString() ?: throw RuntimeException("Format name is null")
        return formats[formatName] ?: throw RuntimeException("Format $formatName not found")
    }

    fun dataFormat(
        formatNameObj: Any?,
        format: String,
    ): Short = dataFormat(formatNameObj) { format }

    fun dataFormat(
        formatNameObj: Any?,
        block: ExcelBuilder.() -> String,
    ): Short {
        val formatName = formatNameObj?.toString() ?: throw RuntimeException("Format name is null")
        logger.debug("Create format with name: {}", formatName)
        val newFormat = workbook.createDataFormat().getFormat(block())
        formats[formatName] = newFormat
        return newFormat
    }

    fun build(block: ExcelBuilder.() -> Unit): SXSSFWorkbook {
        logger.debug("Start build workbook")
        block()
        logger.debug("Finish build workbook")
        return workbook
    }

    fun sheet(
        name: String? = null,
        widthColumns: List<Any?>? = null,
        block: ExcelBuilder.() -> Unit = {},
    ): ExcelBuilder {
        val sheetName = name ?: generateSheetName()
        logger.debug("Create sheet with name: {}", sheetName)
        currentSheet = workbook.getSheet(sheetName) ?: workbook.createSheet(sheetName)
        defaultRowHeight = currentSheet?.defaultRowHeightInPoints ?: -1f

        widthColumns?.forEachIndexed { idx, width ->
            if (width != "default" && width is Number) {
                currentSheet?.setColumnWidth(idx, width.toInt() * 256)
            }
        }

        currentRowIndex = -1
        block()
        return this
    }

    fun row(
        style: String? = null,
        height: Number? = null,
        block: (ExcelBuilder.() -> Unit)? = null,
    ): ExcelBuilder {
        logger.debug("Create row")
        val sheet = currentSheet ?: throw RuntimeException("Current sheet is null")
        currentRowIndex++
        currentRow = sheet.createRow(currentRowIndex)
        currentRowStyle = style?.let { styles[it] }
        if (height != null) {
            currentRow?.heightInPoints = height.toFloat()
        }
        currentCellIndex = -1
        block?.invoke(this)
        currentRow = null
        currentRowStyle = null
        return this
    }

    fun cell(
        style: String? = null,
        rowspan: Int? = null,
        colspan: Int? = null,
        block: (ExcelBuilder.() -> Any?)? = null,
    ): ExcelBuilder {
        logger.debug("Create cell")
        val row = currentRow ?: throw RuntimeException("Current row is null")
        currentCellIndex++
        currentCell = row.getCell(currentCellIndex) ?: row.createCell(currentCellIndex)
        currentCellStyle = style?.let { styles[it] }

        if (rowspan != null || colspan != null) {
            createMergeRegion(rowspan, colspan)
        }

        when {
            currentCellStyle != null -> currentCell?.cellStyle = currentCellStyle
            currentRowStyle != null -> currentCell?.cellStyle = currentRowStyle
        }

        val value = block?.invoke(this)
        when (value) {
            is Number -> currentCell?.setCellValue(value.toDouble())
            is String -> currentCell?.setCellValue(value)
            is Calendar -> currentCell?.setCellValue(value)
            is Date -> currentCell?.setCellValue(value)
            is Boolean -> currentCell?.setCellValue(value)
            null -> currentCell?.setCellValue("")
            else -> {
                logger.warn("Can not determine type of value because value write as string to cell")
                currentCell?.setCellValue(value.toString())
            }
        }

        currentCell = null
        currentCellStyle = null
        return this
    }

    fun skipRow(count: Int = 1): ExcelBuilder {
        repeat(count) {
            row()
        }
        return this
    }

    fun skipCell(count: Int = 1): ExcelBuilder {
        repeat(count) {
            cell()
        }
        return this
    }

    private fun createMergeRegion(
        rowspan: Int?,
        colspan: Int?,
    ) {
        val effectiveRowspan = rowspan ?: 1
        val effectiveColspan = colspan ?: 1
        if (effectiveRowspan == 1 && effectiveColspan == 1) {
            return
        }

        val sheet = currentSheet ?: throw RuntimeException("Current sheet is null")
        sheet.addMergedRegion(
            CellRangeAddress(
                currentRowIndex,
                currentRowIndex + effectiveRowspan - 1,
                currentCellIndex,
                currentCellIndex + effectiveColspan - 1,
            ),
        )

        if (effectiveColspan > 1) {
            val row = currentRow ?: return
            for (i in 1 until effectiveColspan) {
                val additionalCell = row.createCell(currentCellIndex + i)
                when {
                    currentCellStyle != null -> additionalCell.cellStyle = currentCellStyle
                    currentRowStyle != null -> additionalCell.cellStyle = currentRowStyle
                }
            }
        }
    }

    private fun generateSheetName(): String = "Sheet ${workbook.numberOfSheets}"
}
