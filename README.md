# Excel Templater
Templater for excel by example Groovy Swing

Example template:
```java
builder.workbook {
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
```
