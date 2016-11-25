# Excel Builder

## Introduction

Excel Builder позволяет просто и быстро строить Excel таблицы. 
Excel Builder разработан для использования в Java и Groovy.
Excel Builder делает код постороения таблицы простым, прозрачным и читаемым. 

## Install

1. Добавьте репозиторий в gradle проект:

```groovy
repositories {
    maven {
        url "http://repo.serebryanskiy.site/"
    }
}
```

2. Добавьте библиотеку в зависимости проекта:

```groovy
dependencies {
    compile 'maxmoto1702:excel-builder:0.1'
}
```

## Examples

1. Demo types

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
}
```

![Demo types](example-1.png)

2. Demo styles

```groovy
builder.config {
    style("FIRST_STYLE") { CellStyle cellStyle ->
        cellStyle.alignment = CellStyle.ALIGN_CENTER
        cellStyle.borderBottom = CellStyle.BORDER_DASH_DOT
        cellStyle
    }
    style("SECOND_STYLE") { CellStyle cellStyle ->
        cellStyle.alignment = CellStyle.ALIGN_CENTER
        cellStyle
    }
}
builder.build {
    sheet(name: "Demo styles") {
        row(style: "FIRST_STYLE") {
            cell {
                "this cell has row style"
            }
            cell {
                "this cell has row style"
            }
            cell(style: "SECOND_STYLE") {
                "this cell has row style and cell style"
            }
        }
    }
}
```

![Demo types](example-2.png)

3. Demo spans

```groovy
builder.build {
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
}
```

![Demo types](example-3.png)

4. Demo config height and width

```groovy
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
```

![Demo types](example-4.png)

5. Demo dynamic data

```groovy
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
```

![Demo Demo dynamic data](example-5.png)
