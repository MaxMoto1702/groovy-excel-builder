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

    def style(style, Closure closure) {
        log.debug "Create style with name: ${style?.toString()}"
        if (!style) {
            log.error('Style is null')
        }
        closure.delegate = this
        def styleName = style.toString()
        def newStyle = workbook.createCellStyle()
        closure(newStyle)
        styles[styleName] = newStyle
        newStyle
    }

    def font(font) {
        log.debug "Get font with name: $font"
        if (!fonts[font]) {
            log.error "Font ${font?.toString()} not found"
        }
        fonts[font]
    }

    def font(font, Closure closure) {
        log.debug "Create font with name: $font"
        closure.delegate = this
        def newFont = workbook.createFont()
        closure(newFont)
        fonts[font] = newFont
        newFont
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
