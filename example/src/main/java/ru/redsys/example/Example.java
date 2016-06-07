package ru.redsys.example;

import org.apache.poi.ss.usermodel.Workbook;
import ru.redsys.example.model.Book;
import ru.redsys.example.util.ReportType;
import ru.redsys.example.service.ReportService;
import ru.redsys.example.util.Template;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.Map;

public class Example {
    private static final int BOOK_COUNT = 200;
    private static final String FOLDER_PATH = "/tmp";
    private static final File FOLDER = new File(FOLDER_PATH);

    static {
        if (FOLDER.mkdirs()) System.out.println("Folder prepared");
    }

    public static void main(String... args) throws Exception {
        ReportService service = prepareReportService();
        Map data = prepareData();
        Workbook workbook = service.make(ReportType.BOOK_LIST, data);
        save(workbook, "example" + System.currentTimeMillis());
    }

    private static void save(Workbook workbook, String fileName) throws IOException {
        File file = new File(FOLDER, System.currentTimeMillis() + ".xlsx");
        OutputStream stream = new FileOutputStream(file);
        workbook.write(stream);
        stream.flush();
        stream.close();
        System.out.println("Saved to '" + file.getAbsolutePath() + "'");
    }

    private static Map prepareData() {
        Map<String, Object> data = new HashMap<>();
        Collection<Book> books = generateBooks();
        data.put("books", books);
        return data;
    }

    private static Collection<Book> generateBooks() {
        Collection<Book> books = new ArrayList<>();
        for (int bookIndex = 0; bookIndex < BOOK_COUNT; bookIndex++) {
            Book book = generateBook(bookIndex);
            books.add(book);
        }
        return books;
    }

    private static Book generateBook(int bookIndex) {
        Book book = new Book();
        book.setId(bookIndex + 1L);
        book.setTitle("Title #" + book.getId());
        book.setAuthor("Author #" + book.getId());
        book.setAnnotation("Annotation #" + book.getId() + ": Annotation annotation annotation annotation annotation annotation annotation annotation annotation annotation annotation annotation annotation annotation annotation annotation annotation annotation annotation annotation annotation annotation annotation annotation annotation annotation annotation annotation annotation annotation annotation");
        return book;
    }

    private static ReportService prepareReportService() {
        Map<ReportType, Class<? extends Template>> templateClasses = prepareTemplateClasses();
        return new ReportService(templateClasses);
    }

    private static Map<ReportType, Class<? extends Template>> prepareTemplateClasses() {
        Map<ReportType, Class<? extends Template>> templateClasses = new HashMap<>();
        templateClasses.put(ReportType.BOOK_LIST, BookListTemplate.class);
        return templateClasses;
    }
}
