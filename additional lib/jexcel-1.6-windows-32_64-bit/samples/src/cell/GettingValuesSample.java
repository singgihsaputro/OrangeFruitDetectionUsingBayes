/*
 * Copyright (c) 2000-2009 TeamDev Ltd. All rights reserved.
 * TeamDev PROPRIETARY and CONFIDENTIAL.
 * Use is subject to license terms.
 */
package cell;

import com.jniwrapper.win32.jexcel.*;

import java.util.Date;
import java.util.GregorianCalendar;

/**
 * This sample demonstrates how to recieve cell values of different types.
 *
 * @author Vladimir Kondrashchenko
 */
public class GettingValuesSample
{
    private static Worksheet fillSampleData(Worksheet worksheet)
    {
        worksheet.getCell(1, 1).setValue("Apple");
        worksheet.getCell(2, 1).setValue("Orange");
        worksheet.getCell(3, 1).setValue("Strawberry");
        worksheet.getCell(4, 1).setValue("Grapefruit");
        worksheet.getCell(5, 1).setValue("Grape");

        worksheet.getCell(1, 2).setValue(12);
        worksheet.getCell(2, 2).setValue(15);
        worksheet.getCell(3, 2).setValue(10);
        worksheet.getCell(4, 2).setValue(2);
        worksheet.getCell(5, 2).setValue("=SUM(B1:B4)");

        worksheet.getCell(1, 3).setValue(1.1);
        worksheet.getCell(2, 3).setValue(0.23);
        worksheet.getCell(3, 3).setValue(5.1);
        worksheet.getCell(4, 3).setValue(2);
        worksheet.getCell(5, 3).setValue("=SUM(C1:C4)");

        GregorianCalendar calendar = new GregorianCalendar(2000, 0, 12);
        worksheet.getCell(1, 4).setValue(calendar.getTime());

        return worksheet;
    }

    public static void main(String[] args) throws ExcelException
    {
        Application application = new Application();
        GenericWorkbook workbook = application.createWorkbook(null);
        Worksheet worksheet = fillSampleData(workbook.getWorksheet(1));

        Cell cell;

        //Getting cell formula and value
        cell = worksheet.getCell("B5");
        String formula = cell.getFormula();
        long value = cell.getNumber().longValue();
        System.out.println("B5 formula: " + formula);
        System.out.println("B5 value: " + value);

        //Getting string value
        cell = worksheet.getCell("A3");
        String strValue = cell.getString();
        System.out.println("A3 string value: " + strValue);

        //Getting double value
        cell = worksheet.getCell("C2");
        Number numValue = cell.getNumber();
        double doubleValue = numValue.doubleValue();
        System.out.println("C2 double value: " + doubleValue);

        //Getting Date value
        cell = worksheet.getCell("D1");
        Date dateValue = cell.getDate();
        System.out.println("D1 date value: " + dateValue);

        application.close();
    }
}