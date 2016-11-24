package maxmoto1702.excel

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
    }

    def "test styles"() {
        setup:
        def builder = new ExcelBuilder()

        when:
        builder.config {
            style('commonStyle') { cellStyle ->
                cellStyle
            }
            style('customStyle') { cellStyle ->
                cellStyle
            }
        }
        builder.build {
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
        1 == 1
    }

    def "test spans"() {
        setup:
        def builder = new ExcelBuilder()

        when:
        builder.build {
            sheet(name: "Demo spans") {
                row {
                    cell(colspan: 2) {
                        "cell has width 2 columns"
                    }
                    cell { /* dummy */ }
                    cell(rowspan: 2, style: 'wrap', width: 12) {
                        "cell has height 2 rows"
                    }
                    cell(colspan: 2, rowspan: 2, style: 'wrap') {
                        "cell has heigth 2 rows and width 2 columns"
                    }
                }
                // this row not necessary, this row show only dummy-row
                row { /* dummy */ }
            }
        }

        then:
        1 == 1
    }

    def "test config height and width"() {
        setup:
        def builder = new ExcelBuilder()

        when:
        builder.build {
            sheet(name: "Demo config height and width", widthColumns: ['default', 25, 30]) {
                row(height: 10) {
                    cell {
                        "this row has height 30 and default width"
                    }
                }
                row {
                    cell { /* dummy */ }
                    cell(width: 30) {
                        "this column has width 30"
                    }
                    cell(width: 50) {
                        "this column has width 35"
                    }
                }
            }
        }

        then:
        1 == 1
    }

    def "test dynamic data"() {
        setup:
        def builder = new ExcelBuilder()

        when:
        builder.build {
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

                row { /* dummy */ }

                def data2 = [
                        [prop1: "row 1 value 1", prop2: "row 1 value 2"],
                        [prop1: "row 2 value 1", prop2: "row 2 value 2"],
                        [prop1: "row 3 value 1", prop2: "row 3 value 2"],
                        [prop1: null, prop2: "row 4 value 2"]
                ]

                for (def rowData : data2) {
                    row {
                        cell(colspan: rowData.prop ? 1 : 2) {
                            rowData.prop2
                        }
                        if (rowData.prop) {
                            cell {
                                rowData.prop1
                            }
                        }
                    }
                }
            }
        }

        then:
        1 == 1
    }

    def "test some sheets"() {
        setup:
        def builder = new ExcelBuilder()

        when:
        builder.build {
            sheet(name: "Demo sheet 1") {
            }
            sheet(name: "Demo sheet 2") {
            }
            sheet(name: "Demo sheet 3") {
            }
        }

        then:
        1 == 1
    }
}
