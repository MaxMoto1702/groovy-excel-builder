package maxmoto1702.excel

import static maxmoto1702.excel.ExcelBuilderFactory.Style.*

class SampleExcelBuild {
    static void main(String... args) {
        def factory = new ExcelBuilderFactory()
        def builder = factory.getBuilder()
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
                        new Date()
                    }
                    cell {
                        "dyna" + "mic"
                    }
                }
            }
            sheet(name: "Demo styles") {
                row(style: FIRST_STYLE) {
                    cell {
                        "this cell has row style"
                    }
                    cell {
                        "this cell has row style"
                    }
                    cell(style: SECOND_STYLE) {
                        "this cell has row style and cell style"
                    }
                }
            }
            sheet(name: "Demo spans") {
                row {
                    cell(colspan: 2) {
                        "cell has width 2 columns"
                    }
                    cell { /* dummy */ }
                    cell(rowspan: 2, style: WRAP, width: 12) {
                        "cell has height 2 rows"
                    }
                    cell(colspan: 2, rowspan: 2, style: WRAP) {
                        "cell has heigth 2 rows and width 2 columns"
                    }
                }
                // this row not necessary, this row show only dummy-row
                row { /* dummy */ }
            }
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
    }
}
