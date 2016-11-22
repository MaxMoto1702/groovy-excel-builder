package maxmoto1702.excel

import org.apache.poi.ss.usermodel.CellStyle

class ExcelBuilderFactory {
    ExcelBuilder getBuilder() {
        def builder = new ExcelBuilder()
        builder.config {
            style(Style.FIRST_STYLE) { CellStyle cellStyle ->
                cellStyle.alignment = CellStyle.ALIGN_CENTER
                cellStyle.borderBottom = CellStyle.BORDER_DASH_DOT
                cellStyle
            }
        }
        builder
    }

    public static enum Style {
        FIRST_STYLE,
        SECOND_STYLE,
        WRAP
    }
}
