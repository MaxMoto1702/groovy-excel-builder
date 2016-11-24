# Excel Builder

## Install

```groovy
repositories {
    maven {
        url "http://repo.serebryanskiy.site/"
    }
}
dependencies {
    compile 'maxmoto1702:excel-builder:0.1'

    testCompile 'junit:junit:4.12'
}
```

### Configure builder

```groovy
builder.config {
    style(Style.FIRST_STYLE) { CellStyle cellStyle ->
        cellStyle.alignment = CellStyle.ALIGN_CENTER
        cellStyle.borderBottom = CellStyle.BORDER_DASH_DOT
        cellStyle
    }
}
```

### Build excel use closures

#### Example code

```groovy
builder.build {
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
        row(style: rowStyle) {
            cell {
                "this cell has row style"
            }
            cell {
                "this cell has row style"
            }
            cell(style: cellStyle) {
                "this cell has row style and cell style"
            }
        }
    }
    sheet(name: "Demo spans") {
        row {
            cell(colspan: 2) {
                "cell has width 2 columns"
            }
            cell(rowspan: 2, style: wrap, width: 12) {
                "cell has height 2 rows"
            }
            cell(colspan: 2, rowspan: 2, style: wrap) {
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
```

#### Example results
![Demo types](example-1.png)
![Demo styles](example-2.png)
![Demo spans](example-3.png)
![Demo config height and width](example-4.png)
![Demo Demo dynamic data](example-5.png)
