package maxmoto1702.excel

import groovy.util.logging.Slf4j
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.streaming.SXSSFWorkbook

@Slf4j
class ExcelBuilder {

    def workbook = new SXSSFWorkbook()
    def fonts = [:]
    def formats = [:]
    def styles = [:]
    def currentSheet
    def currentRowIndex
    def currentRow
    def currentRowStyle
    def currentCellIndex
    def currentCell
    def currentCellStyle

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// CONFIGURE ///////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    def config(Closure closure) {
        log.debug "Start config builder"
        closure.delegate = this
        closure()
        log.debug "Finish config builder"
        this
    }

    def workbook(Closure closure) {
        workbook = closure()
        this
    }

    def style(styleNameObj, Closure closure) {
        log.debug "Create style with name: ${styleNameObj?.toString()}"
        if (!styleNameObj) {
            log.error('Style name is null')
            throw new RuntimeException('Style name is null')
        }
        closure.delegate = this
        def styleName = styleNameObj.toString()
        def newStyle = workbook.createCellStyle()
        closure(newStyle)
        styles[styleName] = newStyle
        newStyle
    }

    def font(fontNameObj) {
        log.debug "Get font with name: $fontNameObj"
        def fontName = fontNameObj.toString()
        if (!fonts[fontName]) {
            log.error "Font ${fontName} not found"
            throw new RuntimeException("Font ${fontName} not found")
        }
        fonts[fontName]
    }

    def font(fontNameObj, Closure closure) {
        log.debug "Create font with name: $fontNameObj"
        if (!fontNameObj) {
            log.error('Font name is null')
            throw new RuntimeException('Font name is null')
        }
        closure.delegate = this
        def newFont = workbook.createFont()
        closure(newFont)
        def fontName = fontNameObj.toString()
        fonts[fontName] = newFont
        newFont
    }

    def dataFormat(formatNameObj) {
        log.debug "Get format with name: $formatNameObj"
        def formatName = formatNameObj.toString()
        if (!formats[formatName]) {
            log.error "Font $formatName not found"
            throw new RuntimeException("Font $formatName not found")
        }
        formats[formatName]
    }

    def dataFormat(formatNameObj, String format) {
        def closurre = {
            format
        }
        dataFormat(formatNameObj, closurre as Closure)
    }

    def dataFormat(formatNameObj, Closure closure) {
        log.debug "Create format with name: $formatNameObj"
        if (!formatNameObj) {
            log.error('Format name is null')
            throw new RuntimeException('Format name is null')
        }
        def formatName = formatNameObj.toString()
        def newFormat = workbook.createDataFormat().getFormat(closure() as String)
        formats[formatName] = newFormat
        newFormat
    }

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// BUILD ///////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    def build(Closure closure) {
        log.debug "Start build workbook"
        if (!workbook) {
            log.error('Workbook is null')
        }
        closure.delegate = this
        closure()
        log.debug "Finish build workbook"
        workbook
    }

    def sheet(Closure closure) {
        sheet(null, closure)
    }

    def sheet(Map params = null, Closure closure = null) {
        log.debug "Create sheet with params: $params"
        closure.delegate = this
        def sheetName = params?.name ?: generateSheetName()
        currentSheet = workbook.getSheet(sheetName) ?: workbook.createSheet(sheetName)
        if (params?.widthColumns) {
            params?.widthColumns?.eachWithIndex { width, idx ->
                if (width != 'default')
                    currentSheet.setColumnWidth(idx, width * 256)
            }
        }
        currentRowIndex = -1
        closure()
        currentSheet = null
        this
    }

    def row(Closure closure) {
        row(null, closure)
    }

    def row(Map params = null, Closure closure = null) {
        log.debug "Create row with params: $params"
        currentRowIndex++
        currentRow = currentSheet.createRow(currentRowIndex) as Row
        currentRowStyle = styles[params?.style?.toString()]
        if (params?.height)
            currentRow.heightInPoints = params.height
        currentCellIndex = -1
        if (closure) {
            closure.delegate = this
            closure()
        }
        currentRow = null
        currentRowStyle = null
        this
    }

    def cell(Closure closure) {
        cell(null, closure)
    }

    def cell(Map params = null, Closure closure = null) {
        log.debug "Create cell with params: $params"
        currentCellIndex++
        currentCell = currentRow.createCell(currentCellIndex) as Cell
        currentCellStyle = styles[params?.style?.toString()]
        if (params?.rowspan || params?.colspan)
            createMergeRegion(params)
        if (currentCellStyle) {
            log.debug "Apply cell style to current cell"
            currentCell.setCellStyle(currentCellStyle as CellStyle)
        } else if (currentRowStyle) {
            log.debug "Apply row style to current cell"
            currentCell.setCellStyle(currentRowStyle as CellStyle)
        }
        if (closure) {
            closure.delegate = this
            def value = closure()
            log.debug "Set cell value: $value"
            switch (value?.class) {
                case Short:
                case Integer:
                case Long:
                case Float:
                case Double:
                case BigDecimal:
                    currentCell.setCellValue(value as Double)
                    break
                case String:
                    currentCell.setCellValue(value as String)
                    break
                case Calendar:
                    currentCell.setCellValue(value as Calendar)
                    break
                case Date:
                    if (value.time > 0)
                        currentCell.setCellValue(value as Date)
                    else
                        currentCell.setCellValue ""
                    break
                default:
                    log.warn("Can not determine type of value because value write as string to cell")
                    currentCell.setCellValue "$value"
            }
        }
        currentCell = null
        currentCellStyle = null
        this
    }

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// UTILS ///////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    def skipRow(count = 1) {
        (1..count).each {
            row()
        }
        this
    }

    def skipCell(count = 1) {
        (1..count).each {
            cell()
        }
        this
    }

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// PRIVATE UTILS ///////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    private def createMergeRegion(params) {
        def rowspan = params?.rowspan ?: 1
        def colspan = params?.colspan ?: 1
        log.debug "Create merge region : $rowspan, $colspan"
        if (rowspan == 1 && colspan == 1) {
            log.warn "Merge region $rowspan x $colspan not create!"
            return
        }
        currentSheet.addMergedRegion(new CellRangeAddress(
                currentRowIndex as Integer,
                currentRowIndex + rowspan - 1 as Integer,
                currentCellIndex as Integer,
                currentCellIndex + colspan - 1 as Integer
        ))
    }

    private String generateSheetName() {
        def generatedSheetName = "Sheet $workbook.numberOfSheets"
        log.debug "Generated sheet name: $generatedSheetName"
        generatedSheetName
    }
}
