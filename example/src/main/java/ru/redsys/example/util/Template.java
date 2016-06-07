package ru.redsys.example.util;

import ru.redsys.exceltemplater.core.WorkbookBuilder;

public interface Template {
    byte[] RGB_BLUE = new byte[]{0, (byte) 0xC5, (byte) 0xD9, (byte) 0xF1};
    WorkbookBuilder build(WorkbookBuilder builder, Object data);
}
