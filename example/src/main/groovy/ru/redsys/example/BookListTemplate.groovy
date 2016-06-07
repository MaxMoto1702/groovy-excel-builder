package ru.redsys.example

import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.xssf.streaming.SXSSFWorkbook
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFColor
import ru.redsys.example.model.Book
import ru.redsys.example.util.Template
import ru.redsys.exceltemplater.core.WorkbookBuilder

import static ru.redsys.example.style.AlignStyle.*
import static ru.redsys.example.style.VerticalAlignStyle.*
import static ru.redsys.example.style.BackgroundStyle.*
import static ru.redsys.example.style.BorderStyle.*

class BookListTemplate implements Template {
    @Override
    WorkbookBuilder build(WorkbookBuilder builder, Object data) {
        builder.workbook {
            styles {
                style(border: SOLID) {
                    if (workbook instanceof SXSSFWorkbook) {
                        XSSFCellStyle style = (XSSFCellStyle) workbook.createCellStyle();
                        style.setBorderBottom(CellStyle.BORDER_THIN);
                        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
                        style.setBorderLeft(CellStyle.BORDER_THIN);
                        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
                        style.setBorderRight(CellStyle.BORDER_THIN);
                        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
                        style.setBorderTop(CellStyle.BORDER_THIN);
                        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
                        mixes.put(Arrays.asList(SOLID), style);
                    }
                }
                style(backgroundColor: BLUE) {
                    if (workbook instanceof SXSSFWorkbook) {
                        XSSFCellStyle style = (XSSFCellStyle) workbook.createCellStyle();
                        XSSFColor color = new XSSFColor(RGB_BLUE);
                        style.setFillForegroundColor(color);
                        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                        mixes.put(Arrays.asList(BLUE), style);
                    }
                }
                style(align: CENTER) {
                    if (workbook instanceof SXSSFWorkbook) {
                        XSSFCellStyle style = (XSSFCellStyle) workbook.createCellStyle();
                        style.setAlignment(CellStyle.ALIGN_CENTER);
                        mixes.put(Arrays.asList(CENTER), style);
                    }
                }
                style(verticalAlign: MIDDLE) {
                    if (workbook instanceof SXSSFWorkbook) {
                        XSSFCellStyle style = (XSSFCellStyle) workbook.createCellStyle();
                        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
                        mixes.put(Arrays.asList(MIDDLE), style);
                    }
                }
            }
            sheet("Books") {
                // Title
                row(height: 25) {
                    cell(colspan: 4, rowspan: 2, border: SOLID, align: CENTER, verticalAlign: MIDDLE) {
                        "List of books"
                    }
                }
                row()
                // Headers of book list
                row {
                    cell {
                        "#"
                    }
                    cell {
                        "Title"
                    }
                    cell {
                        "Author"
                    }
                    cell {
                        "Annotation"
                    }
                }
                // Books
                data.books.each { Book book ->
                    row {
                        cell {
                            book.id
                        }
                        cell {
                            book.title
                        }
                        cell {
                            book.author
                        }
                        cell {
                            book.annotation
                        }
                    }
                }
                // Footer of book list
                row {
                    cell {
                        "Count of books"
                    }
                    cell {
                        data.books?.size()
                    }
                }
            }
            sheet("Page 2") {
                row {
                    cell {
                        "test"
                    }
                }
            }
        }
    }
}
