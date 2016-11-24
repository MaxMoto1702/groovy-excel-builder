package maxmoto1702.excel

import groovy.util.logging.Slf4j
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellStyle
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
        closure.delegate = this
        closure()
        this
    }

    def workbook(Closure closure) {
        workbook = closure()
        this
    }

    def style(style, Closure closure) {
        if (!style) {
            log.error('Style is null')
        }
        closure.delegate = this
        def styleName = style.toString()
        styles[styleName] = closure(workbook.createCellStyle())
    }

    def font(String name, Closure closure) {
        closure.delegate = this
        fonts[name] = closure(workbook.createFont())
    }

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// BUILD ///////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    def build(Closure closure) {
        if (!workbook) {
            log.error('Workbook is null')
        }
        closure.delegate = this
        closure()
        workbook
    }

    def sheet(Closure closure) {
        sheet(null, closure)
    }

    def sheet(Map params = null, Closure closure = null) {
        closure.delegate = this
        def sheetName = params?.name ?: generateSheetName()
        log.debug "Create sheet with name '$sheetName"
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
        currentRowIndex++
        currentRow = currentSheet.createRow(currentRowIndex) as Row
//        currentRowStyle = params?.style && styles[params?.style?.toString()]
//        if (params?.style && styles[params?.style?.toString()])
//            currentRow.rowStyle = styles[params?.style?.toString()] as CellStyle
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
        currentCellIndex++
        currentCell = currentRow.createCell(currentCellIndex) as Cell
        currentCellStyle = params?.style && styles[params?.style?.toString()]
        if (params?.rowspan || params?.colspan)
            createMergeRegion(params.rowspan, params.colspan)
        if (currentCellStyle)
            currentCell.cellStyle = currentCellStyle as CellStyle
        else if(currentRowStyle)
            currentCell.cellStyle = currentRowStyle as CellStyle
        if (closure) {
            closure.delegate = this
            def v = closure()
            currentCell.setCellValue "$v"
        }
        currentCell = null
        currentCellStyle = null
        this
    }

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// UTILS ///////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    private def createMergeRegion(rowspan, colspan) {
        currentSheet.addMergedRegion(new CellRangeAddress(
                currentRowIndex as Integer,
                currentRowIndex + (rowspan ?: 1) - 1 as Integer,
                currentCellIndex as Integer,
                currentCellIndex + (colspan ?: 1) - 1 as Integer
        ))
    }

    private String generateSheetName() {
        "Sheet $workbook.numberOfSheets"
    }
}
