package net.serebryansky.excel_builder

import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.xssf.streaming.SXSSFWorkbook
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.assertj.core.api.Assertions.assertThat
import org.junit.jupiter.api.Test
import java.util.Date

class ExcelBuilderTest {
    @Test
    fun `test types`() {
        val builder = ExcelBuilder()
        val date = Date()

        val workbook =
            builder.build {
                sheet(name = "Demo types") {
                    row {
                        cell { "string value" }
                        cell { 1 }
                        cell { date }
                        cell { "dyna" + "mic" }
                    }
                }
            }

        val row = workbook.getSheet("Demo types").getRow(0)
        assertThat(row.getCell(0).stringCellValue).isEqualTo("string value")
        assertThat(row.getCell(1).numericCellValue).isEqualTo(1.0)
        assertThat(row.getCell(2).dateCellValue).isEqualTo(date)
        assertThat(row.getCell(3).stringCellValue).isEqualTo("dynamic")
    }

    @Test
    fun `test styles`() {
        val builder = ExcelBuilder()

        builder.config {
            style("commonStyle") { cellStyle ->
                cellStyle.alignment = HorizontalAlignment.CENTER
                cellStyle.borderBottom = BorderStyle.DASH_DOT
            }
            style("customStyle") { cellStyle ->
                cellStyle.alignment = HorizontalAlignment.LEFT
                cellStyle.borderTop = BorderStyle.DOUBLE
            }
        }

        val workbook =
            builder.build {
                sheet(name = "Demo styles") {
                    row(style = "commonStyle") {
                        cell { "this cell has row style" }
                        cell { "this cell has row style" }
                        cell(style = "customStyle") { "this cell has common style and custom style" }
                    }
                }
            }

        val sheet = workbook.getSheet("Demo styles")
        val row = sheet.getRow(0)
        assertThat(row.getCell(0).cellStyle.alignment).isEqualTo(HorizontalAlignment.CENTER)
        assertThat(row.getCell(0).cellStyle.borderBottom).isEqualTo(BorderStyle.DASH_DOT)
        assertThat(row.getCell(1).cellStyle.alignment).isEqualTo(HorizontalAlignment.CENTER)
        assertThat(row.getCell(1).cellStyle.borderBottom).isEqualTo(BorderStyle.DASH_DOT)
        assertThat(row.getCell(2).cellStyle.alignment).isEqualTo(HorizontalAlignment.LEFT)
        assertThat(row.getCell(2).cellStyle.borderTop).isEqualTo(BorderStyle.DOUBLE)
    }

    @Test
    fun `test fonts`() {
        val builder = ExcelBuilder()

        builder.config {
            font("waterfall configured font") { font ->
                font.fontName = "Arial"
            }
            style("style with waterfall configured font") { cellStyle ->
                cellStyle.setFont(font("waterfall configured font"))
            }
            style("style with included configured font") { cellStyle ->
                cellStyle.setFont(
                    font("included configured font") { font ->
                        font.fontName = "Arial"
                    },
                )
            }
        }

        val workbook =
            builder.build {
                sheet {
                    row {
                        cell(style = "style with waterfall configured font") { "test waterfall configured font" }
                        cell(style = "style with included configured font") { "test included configured font" }
                    }
                }
            }

        val row = workbook.getSheetAt(0).getRow(0)
        assertThat(row.getCell(0).stringCellValue).isEqualTo("test waterfall configured font")
        val style0 = row.getCell(0).cellStyle as XSSFCellStyle
        assertThat(style0.font.fontName).isEqualTo("Arial")
        assertThat(row.getCell(1).stringCellValue).isEqualTo("test included configured font")
        val style1 = row.getCell(1).cellStyle as XSSFCellStyle
        assertThat(style1.font.fontName).isEqualTo("Arial")
    }

    @Test
    fun `test formats`() {
        val builder = ExcelBuilder()

        builder.config {
            dataFormat("1", "# ##0.00")
            style("1") { cellStyle ->
                cellStyle.dataFormat = dataFormat("1")
            }

            style("2") { cellStyle ->
                cellStyle.dataFormat = dataFormat("2", "# ### ##0.00")
            }

            dataFormat("3") { "# ### ### ##0.00" }
            style("3") { cellStyle ->
                cellStyle.dataFormat = dataFormat("3")
            }

            style("4") { cellStyle ->
                cellStyle.dataFormat = dataFormat("4") { "# ### ### ### ##0.00" }
            }
        }

        val workbook =
            builder.build {
                sheet {
                    row {
                        cell(style = "1") { "test style with waterfall configured format (string)" }
                        cell(style = "2") { "test style with included configured format (string)" }
                        cell(style = "3") { "test style with waterfall configured format (closure)" }
                        cell(style = "4") { "test style with included configured format (closure)" }
                    }
                }
            }

        val row = workbook.getSheetAt(0).getRow(0)
        assertThat(row.getCell(0).cellStyle.dataFormatString).isEqualTo("# ##0.00")
        assertThat(row.getCell(1).cellStyle.dataFormatString).isEqualTo("# ### ##0.00")
        assertThat(row.getCell(2).cellStyle.dataFormatString).isEqualTo("# ### ### ##0.00")
        assertThat(row.getCell(3).cellStyle.dataFormatString).isEqualTo("# ### ### ### ##0.00")
    }

    @Test
    fun `test spans`() {
        val builder = ExcelBuilder()

        val workbook =
            builder.build {
                sheet(name = "Demo spans") {
                    row {
                        cell(colspan = 2) { "cell has width 2 columns" }
                        cell()
                        cell(rowspan = 2, style = "wrap") { "cell has height 2 rows" }
                        cell(colspan = 2, rowspan = 2, style = "wrap") { "cell has heigth 2 rows and width 2 columns" }
                    }
                    row()
                }
            }

        val sheet = workbook.getSheet("Demo spans")
        assertThat(sheet.getMergedRegion(0).formatAsString()).isEqualTo("A1:B1")
        assertThat(sheet.getMergedRegion(1).formatAsString()).isEqualTo("C1:C2")
        assertThat(sheet.getMergedRegion(2).formatAsString()).isEqualTo("D1:E2")
    }

    @Test
    fun `test config height and width`() {
        val builder = ExcelBuilder()

        val workbook =
            builder.build {
                sheet(name = "Demo config height and width", widthColumns = listOf("default", 25, 30)) {
                    row(height = 10) {
                        cell { "this row has height 30 and default width" }
                    }
                    row {
                        cell()
                        cell { "this column has width 30" }
                        cell { "this column has width 35" }
                    }
                    row(height = ExcelBuilder.defaultRowHeight) {
                        cell()
                    }
                }
            }

        val sheet = workbook.getSheet("Demo config height and width")
        assertThat(sheet.getColumnWidth(0) / 256).isEqualTo(sheet.defaultColumnWidth)
        assertThat(sheet.getColumnWidth(1) / 256).isEqualTo(25)
        assertThat(sheet.getColumnWidth(2) / 256).isEqualTo(30)
        assertThat(sheet.getRow(0).height / 20).isEqualTo(10)
        assertThat(sheet.getRow(1).height).isEqualTo(sheet.defaultRowHeight)
        assertThat(workbook.getSheetAt(0).getRow(2).height).isEqualTo(workbook.getSheetAt(0).defaultRowHeight)
    }

    @Test
    fun `test dynamic data`() {
        val builder = ExcelBuilder()

        val workbook =
            builder.build {
                sheet(name = "Demo dynamic data", widthColumns = listOf(12, 12)) {
                    val data1 =
                        listOf(
                            listOf("value 1", "value 2"),
                            listOf("value 3", "value 4"),
                            listOf("value 5", "value 6"),
                        )

                    data1.forEach { rowData ->
                        row {
                            rowData.forEach { cellData ->
                                cell { cellData }
                            }
                        }
                    }

                    row()

                    val data2 =
                        listOf(
                            mapOf("prop1" to "row 1 value 1", "prop2" to "row 1 value 2"),
                            mapOf("prop1" to "row 2 value 1", "prop2" to "row 2 value 2"),
                            mapOf("prop1" to "row 3 value 1", "prop2" to "row 3 value 2"),
                            mapOf("prop2" to "row 4 value 2"),
                        )

                    data2.forEach { rowData ->
                        row {
                            val prop1 = rowData["prop1"]
                            val prop2 = rowData["prop2"]
                            cell(colspan = if (prop1 != null) 1 else 2) { prop2 }
                            if (prop1 != null) {
                                cell { prop1 }
                            }
                        }
                    }
                }
            }

        val sheet = workbook.getSheet("Demo dynamic data")
        assertThat(sheet.getRow(0).getCell(0).stringCellValue).isEqualTo("value 1")
        assertThat(sheet.getRow(0).getCell(1).stringCellValue).isEqualTo("value 2")
        assertThat(sheet.getRow(7).getCell(0).stringCellValue).isEqualTo("row 4 value 2")
        assertThat(sheet.getMergedRegion(0).formatAsString()).isEqualTo("A8:B8")
    }

    @Test
    fun `test several sheets`() {
        val builder = ExcelBuilder()

        val workbook =
            builder.build {
                sheet { }
                sheet { }
                sheet { }
            }

        assertThat(workbook.getSheetAt(0)).isNotNull
        assertThat(workbook.getSheetAt(1)).isNotNull
        assertThat(workbook.getSheetAt(2)).isNotNull
    }

    @Test
    fun `test skipping`() {
        val builder = ExcelBuilder()

        val workbook =
            builder.build {
                sheet {
                    row {
                        cell { "A" }
                        skipCell()
                        cell { "B" }
                        skipCell(2)
                        cell { "C" }
                    }
                    skipRow()
                    row {
                        cell { "D" }
                    }
                    skipRow(2)
                    row {
                        cell { "E" }
                    }
                }
            }

        val sheet = workbook.getSheetAt(0)
        assertThat(sheet.getRow(0).getCell(0).stringCellValue).isEqualTo("A")
        assertThat(sheet.getRow(0).getCell(2).stringCellValue).isEqualTo("B")
        assertThat(sheet.getRow(0).getCell(5).stringCellValue).isEqualTo("C")
        assertThat(sheet.getRow(2).getCell(0).stringCellValue).isEqualTo("D")
        assertThat(sheet.getRow(5).getCell(0).stringCellValue).isEqualTo("E")
    }

    @Test
    fun `test workbook function`() {
        val externalWorkbook = SXSSFWorkbook()
        val builder = ExcelBuilder()

        builder.workbook { externalWorkbook }

        val workbook =
            builder.build {
                sheet(name = "External workbook sheet") {
                    row {
                        cell { "from external workbook" }
                    }
                }
            }

        assertThat(workbook).isSameAs(externalWorkbook)
        assertThat(workbook.getSheet("External workbook sheet").getRow(0).getCell(0).stringCellValue)
            .isEqualTo("from external workbook")
    }
}
