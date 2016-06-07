package ru.redsys.example.service;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import ru.redsys.example.util.ReportType;
import ru.redsys.example.util.Template;
import ru.redsys.exceltemplater.core.WorkbookBuilder;

import java.util.Map;

public class ReportService {
    Map<ReportType, Class<? extends Template>> templateClasses;

    public ReportService(Map<ReportType, Class<? extends Template>> templateClasses) {
        this.templateClasses = templateClasses;
    }

    public Workbook make(ReportType reportType, Object data) throws ClassNotFoundException, InstantiationException, IllegalAccessException {
        Template template = getTemplate(reportType);
        WorkbookBuilder builder = new WorkbookBuilder(new SXSSFWorkbook());
        return template.build(builder, data).getWorkbook();
    }

    public Map<ReportType, Class<? extends Template>> getTemplateClasses() {
        return templateClasses;
    }

    public void setTemplateClasses(Map<ReportType, Class<? extends Template>> templateClasses) {
        this.templateClasses = templateClasses;
    }

    private Template getTemplate(ReportType reportType) throws ClassNotFoundException, IllegalAccessException, InstantiationException {
        if (!templateClasses.containsKey(reportType)) throw new ClassNotFoundException("Not found template class for report type '" + reportType + "'");
        Class<? extends Template> templateClass = templateClasses.get(reportType);
        return templateClass.newInstance();
    }
}
