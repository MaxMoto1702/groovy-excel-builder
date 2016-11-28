package maxmoto1702.excel

import org.apache.poi.ss.usermodel.CellStyle
import spock.lang.*

class ExcelBuilderTests extends Specification {

    def "test types"() {
        setup:
        def builder = new ExcelBuilder()
        def date = new Date()

        when:
        def workbook = builder.build {
            sheet(name: "Demo types") {
                row {
                    cell {
                        "string value"
                    }
                    cell {
                        // integer value
                        1
                    }
                    cell {
                        // date value
                        date
                    }
                    cell {
                        "dyna" + "mic"
                    }
                }
            }
        }

        then:
        workbook.getSheet("Demo types") != null
        workbook.getSheet("Demo types").getRow(0) != null
        workbook.getSheet("Demo types").getRow(0).getCell(0)?.stringCellValue == "string value"
        workbook.getSheet("Demo types").getRow(0).getCell(1)?.numericCellValue == 1
        workbook.getSheet("Demo types").getRow(0).getCell(2)?.dateCellValue == date
        workbook.getSheet("Demo types").getRow(0).getCell(3)?.stringCellValue == "dynamic"
    }

    def "test styles"() {
        setup:
        def builder = new ExcelBuilder()

        when:
        builder.config {
            style('commonStyle') { cellStyle ->
                cellStyle.alignment = CellStyle.ALIGN_CENTER
                cellStyle.borderBottom = CellStyle.BORDER_DASH_DOT
                cellStyle
            }
            style('customStyle') { cellStyle ->
                cellStyle.alignment = CellStyle.ALIGN_LEFT
                cellStyle.borderTop = CellStyle.BORDER_DOUBLE
                cellStyle
            }
        }
        def workbook = builder.build {
            sheet(name: "Demo styles") {
                row(style: 'commonStyle') {
                    cell {
                        "this cell has row style"
                    }
                    cell {
                        "this cell has row style"
                    }
                    cell(style: 'customStyle') {
                        "this cell has common style and custom style"
                    }
                }
            }
        }

        then:
        workbook.getSheet("Demo styles") != null
        workbook.getSheet("Demo styles").getRow(0) != null
        workbook.getSheet("Demo styles").getRow(0).getCell(0).getCellStyle().alignment == CellStyle.ALIGN_CENTER
        workbook.getSheet("Demo styles").getRow(0).getCell(0).getCellStyle().borderBottom == CellStyle.BORDER_DASH_DOT
        workbook.getSheet("Demo styles").getRow(0).getCell(1).cellStyle.alignment == CellStyle.ALIGN_CENTER
        workbook.getSheet("Demo styles").getRow(0).getCell(1).cellStyle.borderBottom == CellStyle.BORDER_DASH_DOT
        workbook.getSheet("Demo styles").getRow(0).getCell(2).cellStyle.alignment == CellStyle.ALIGN_LEFT
        workbook.getSheet("Demo styles").getRow(0).getCell(2).cellStyle.borderTop == CellStyle.BORDER_DOUBLE
    }

    def "test spans"() {
        setup:
        def builder = new ExcelBuilder()

        when:
        def workbook = builder.build {
            sheet(name: "Demo spans") {
                row {
                    // A1:B1
                    cell(colspan: 2) {
                        "cell has width 2 columns"
                    }
                    cell()
                    // C1:C2
                    cell(rowspan: 2, style: 'wrap', width: 12) {
                        "cell has height 2 rows"
                    }
                    // D1:E2
                    cell(colspan: 2, rowspan: 2, style: 'wrap') {
                        "cell has heigth 2 rows and width 2 columns"
                    }
                }
                // this row not necessary, this row show only dummy-row
                row()
            }
        }

        then:
        workbook.getSheet("Demo spans") != null
        workbook.getSheet("Demo spans").getRow(0) != null
        workbook.getSheet("Demo spans").getRow(0).getCell(0) != null
        workbook.getSheet("Demo spans").getRow(0).getCell(1) != null
        workbook.getSheet("Demo spans").getRow(0).getCell(2) != null
        workbook.getSheet("Demo spans").getRow(0).getCell(3) != null
        workbook.getSheet("Demo spans").getRow(1) != null
        workbook.getSheet("Demo spans").getRow(0).getCell(0) != null
        workbook.getSheet("Demo spans").getMergedRegion(0) != null
        workbook.getSheet("Demo spans").getMergedRegion(0).formatAsString() == "A1:B1"
        workbook.getSheet("Demo spans").getMergedRegion(1) != null
        workbook.getSheet("Demo spans").getMergedRegion(1).formatAsString() == "C1:C2"
        workbook.getSheet("Demo spans").getMergedRegion(2) != null
        workbook.getSheet("Demo spans").getMergedRegion(2).formatAsString() == "D1:E2"
    }

    def "test config height and width"() {
        setup:
        def builder = new ExcelBuilder()

        when:
        def workbook = builder.build {
            sheet(name: "Demo config height and width", widthColumns: ['default', 25, 30]) {
                row(height: 10) {
                    cell {
                        "this row has height 30 and default width"
                    }
                }
                row {
                    cell()
                    cell {
                        "this column has width 30"
                    }
                    cell {
                        "this column has width 35"
                    }
                }
            }
        }

        then:
        workbook.getSheet("Demo config height and width") != null
        workbook.getSheet("Demo config height and width").getColumnWidth(0) / 256 as Integer == workbook.getSheet("Demo config height and width").defaultColumnWidth
        workbook.getSheet("Demo config height and width").getColumnWidth(1) / 256 as Integer == 25
        workbook.getSheet("Demo config height and width").getColumnWidth(2) / 256 as Integer == 30
        workbook.getSheet("Demo config height and width").getRow(0) != null
        workbook.getSheet("Demo config height and width").getRow(0).height / 20 as Integer == 10
        workbook.getSheet("Demo config height and width").getRow(0).getCell(0) != null
        workbook.getSheet("Demo config height and width").getRow(0).getCell(0) != null
        workbook.getSheet("Demo config height and width").getRow(1) != null
        workbook.getSheet("Demo config height and width").getRow(1).height == workbook.getSheet("Demo config height and width").defaultRowHeight
        workbook.getSheet("Demo config height and width").getRow(1).getCell(0) != null
        workbook.getSheet("Demo config height and width").getRow(1).getCell(1) != null
        workbook.getSheet("Demo config height and width").getRow(1).getCell(2) != null
    }

    def "test dynamic data"() {
        setup:
        def builder = new ExcelBuilder()

        when:
        def workbook = builder.build {
            sheet(name: "Demo dynamic data", widthColumns: [12, 12]) {
                def data1 = [
                        ["value 1", "value 2"],
                        ["value 3", "value 4"],
                        ["value 5", "value 6"]
                ]

                for (def rowData : data1) {
                    row {
                        for (def cellData : rowData) {
                            cell {
                                cellData
                            }
                        }
                    }
                }

                row()

                def data2 = [
                        [prop1: "row 1 value 1", prop2: "row 1 value 2"],
                        [prop1: "row 2 value 1", prop2: "row 2 value 2"],
                        [prop1: "row 3 value 1", prop2: "row 3 value 2"],
                        // For this data row will create merge region A8:B8
                        [prop2: "row 4 value 2"]
                ]

                for (def rowData : data2) {
                    row {
                        cell(colspan: rowData.prop1 ? 1 : 2) {
                            rowData.prop2
                        }
                        if (rowData.prop1) {
                            cell {
                                rowData.prop1
                            }
                        }
                    }
                }
            }
        }

        then:
        workbook.getSheet("Demo dynamic data") != null
        workbook.getSheet("Demo dynamic data").getRow(0) != null
        workbook.getSheet("Demo dynamic data").getRow(0).getCell(0) != null
        workbook.getSheet("Demo dynamic data").getRow(0).getCell(0).stringCellValue == "value 1"
        workbook.getSheet("Demo dynamic data").getRow(0).getCell(1) != null
        workbook.getSheet("Demo dynamic data").getRow(0).getCell(1).stringCellValue == "value 2"
        workbook.getSheet("Demo dynamic data").getRow(1) != null
        workbook.getSheet("Demo dynamic data").getRow(1).getCell(0) != null
        workbook.getSheet("Demo dynamic data").getRow(1).getCell(0).stringCellValue == "value 3"
        workbook.getSheet("Demo dynamic data").getRow(1).getCell(1) != null
        workbook.getSheet("Demo dynamic data").getRow(1).getCell(1).stringCellValue == "value 4"
        workbook.getSheet("Demo dynamic data").getRow(2) != null
        workbook.getSheet("Demo dynamic data").getRow(2).getCell(0) != null
        workbook.getSheet("Demo dynamic data").getRow(2).getCell(0).stringCellValue == "value 5"
        workbook.getSheet("Demo dynamic data").getRow(2).getCell(1) != null
        workbook.getSheet("Demo dynamic data").getRow(2).getCell(1).stringCellValue == "value 6"
        workbook.getSheet("Demo dynamic data").getRow(3) != null
        workbook.getSheet("Demo dynamic data").getRow(4) != null
        workbook.getSheet("Demo dynamic data").getRow(4).getCell(0) != null
        workbook.getSheet("Demo dynamic data").getRow(4).getCell(0).stringCellValue == "row 1 value 2"
        workbook.getSheet("Demo dynamic data").getRow(4).getCell(1) != null
        workbook.getSheet("Demo dynamic data").getRow(4).getCell(1).stringCellValue == "row 1 value 1"
        workbook.getSheet("Demo dynamic data").getRow(5) != null
        workbook.getSheet("Demo dynamic data").getRow(5).getCell(0) != null
        workbook.getSheet("Demo dynamic data").getRow(5).getCell(0).stringCellValue == "row 2 value 2"
        workbook.getSheet("Demo dynamic data").getRow(5).getCell(1) != null
        workbook.getSheet("Demo dynamic data").getRow(5).getCell(1).stringCellValue == "row 2 value 1"
        workbook.getSheet("Demo dynamic data").getRow(6) != null
        workbook.getSheet("Demo dynamic data").getRow(6).getCell(0) != null
        workbook.getSheet("Demo dynamic data").getRow(6).getCell(0).stringCellValue == "row 3 value 2"
        workbook.getSheet("Demo dynamic data").getRow(6).getCell(1) != null
        workbook.getSheet("Demo dynamic data").getRow(6).getCell(1).stringCellValue == "row 3 value 1"
        workbook.getSheet("Demo dynamic data").getRow(7) != null
        workbook.getSheet("Demo dynamic data").getRow(7).getCell(0) != null
        workbook.getSheet("Demo dynamic data").getRow(7).getCell(0).stringCellValue == "row 4 value 2"
        workbook.getSheet("Demo dynamic data").getMergedRegion(0) != null
        workbook.getSheet("Demo dynamic data").getMergedRegion(0).formatAsString() == "A8:B8"
    }

    def "test several sheets"() {
        setup:
        def builder = new ExcelBuilder()

        when:
        def workbook = builder.build {
            sheet() {
            }
            sheet() {
            }
            sheet() {
            }
        }

        then:
        workbook.getSheetAt(0) != null
        workbook.getSheetAt(1) != null
        workbook.getSheetAt(2) != null
    }

    def "test skipping"() {
        setup:
        def builder = new ExcelBuilder()

        when:
        def workbook = builder.build {
            sheet {
                row {
                    cell {
                        'A'
                    }
                    skipCell()
                    cell {
                        'B'
                    }
                    skipCell(2)
                    cell {
                        'C'
                    }
                }
                skipRow()
                row {
                    cell {
                        'D'
                    }
                }
                skipRow(2)
                row {
                    cell {
                        'E'
                    }
                }
            }
        }

        then:
        workbook.getSheetAt(0) != null
        workbook.getSheetAt(0).getRow(0) != null
        workbook.getSheetAt(0).getRow(0).getCell(0) != null
        workbook.getSheetAt(0).getRow(0).getCell(0).stringCellValue == 'A'
        workbook.getSheetAt(0).getRow(0).getCell(1) != null
        workbook.getSheetAt(0).getRow(0).getCell(2) != null
        workbook.getSheetAt(0).getRow(0).getCell(2).stringCellValue == 'B'
        workbook.getSheetAt(0).getRow(0).getCell(3) != null
        workbook.getSheetAt(0).getRow(0).getCell(4) != null
        workbook.getSheetAt(0).getRow(0).getCell(5) != null
        workbook.getSheetAt(0).getRow(0).getCell(5).stringCellValue == 'C'
        workbook.getSheetAt(0).getRow(1) != null
        workbook.getSheetAt(0).getRow(2) != null
        workbook.getSheetAt(0).getRow(2).getCell(0) != null
        workbook.getSheetAt(0).getRow(2).getCell(0).stringCellValue == 'D'
        workbook.getSheetAt(0).getRow(3) != null
        workbook.getSheetAt(0).getRow(4) != null
        workbook.getSheetAt(0).getRow(5) != null
        workbook.getSheetAt(0).getRow(5).getCell(0) != null
        workbook.getSheetAt(0).getRow(5).getCell(0).stringCellValue == 'E'
    }

}
