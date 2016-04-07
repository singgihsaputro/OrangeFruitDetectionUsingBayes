/*
 * Copyright (c) 2000-2009 TeamDev Ltd. All rights reserved.
 * TeamDev PROPRIETARY and CONFIDENTIAL.
 * Use is subject to license terms.
 */
package cell;

import com.jniwrapper.win32.jexcel.*;

import java.util.Date;

/**
 * This sample demonstrates how to set cell values of different types.
 *
 * @author Vladimir Kondrashchenko
 */
public class SettingValuesSample
{
    public static void main(String[] args) throws ExcelException
    {
        Application application = new Application();

        GenericWorkbook workbook = application.createWorkbook(null);

        Worksheet worksheet = workbook.getWorksheet(1);

        Cell cell = worksheet.getCell("A1");
        //Setting string value
        cell.setValue("String value");

        cell = worksheet.getCell("A2");
        //Setting long value
        cell.setValue(220);

        cell = worksheet.getCell("A3");
        //Setting double value
        cell.setValue(122.1);

        cell = worksheet.getCell("A4");
        //Setting Date value
        cell.setValue(new Date());

        //Setting formula
        cell.setValue("=SUM(A1:B12)");

        application.close();
    }
}