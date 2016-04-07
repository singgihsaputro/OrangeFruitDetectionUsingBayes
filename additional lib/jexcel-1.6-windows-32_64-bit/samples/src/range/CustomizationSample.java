/*
 * Copyright (c) 2000-2009 TeamDev Ltd. All rights reserved.
 * TeamDev PROPRIETARY and CONFIDENTIAL.
 * Use is subject to license terms.
 */
package range;

import com.jniwrapper.win32.jexcel.*;
import com.jniwrapper.win32.jexcel.format.TextAlignment;
import com.jniwrapper.win32.jexcel.format.TextOrientation;

/**
 * @author Vladimir Kondrashchenko
 */
public class CustomizationSample
{
    private static Worksheet fillSampleData(Worksheet worksheet)
    {
        worksheet.getCell(1, 1).setValue("Apple");
        worksheet.getCell(2, 1).setValue("Orange");
        worksheet.getCell(3, 1).setValue("Strawberry");
        worksheet.getCell(4, 1).setValue("Grapefruit");
        worksheet.getCell(5, 1).setValue("Apple");

        worksheet.getCell(1, 2).setValue(12);
        worksheet.getCell(2, 2).setValue(15);
        worksheet.getCell(3, 2).setValue(10);
        worksheet.getCell(4, 2).setValue(2);
        worksheet.getCell(5, 2).setValue(100);

        worksheet.getCell(1, 3).setValue(1.1);
        worksheet.getCell(2, 3).setValue(0.23);
        worksheet.getCell(3, 3).setValue(5.1);
        worksheet.getCell(4, 3).setValue(2);
        worksheet.getCell(5, 3).setValue(0.01);

        return worksheet;
    }

    public static void main(String[] args) throws ExcelException
    {
        Application application = new Application();
        GenericWorkbook workbook = application.createWorkbook(null);
        Worksheet worksheet = fillSampleData(workbook.getWorksheet(1));

        //Getting a range number format
        Range range = worksheet.getRange("A1:A3");
        String numberFormat = range.getNumberFormat();
        System.out.println("A1:A3 number format is " + numberFormat);

        //Setting custom number format
        String newNumberFormat = "0,00%";
        range.setNumberFormat(newNumberFormat);
        System.out.println("A1:A3 new number format is " + range.getNumberFormat());

        //Setting up custom horizontal text alignment
        range.setHorizontalAlignment(TextAlignment.RIGHT);

        //Verifing horizontal text alignment
        if (range.getHorizontalAlignment().equals(TextAlignment.RIGHT))
        {
            System.out.println("A1:A3 range: new horizontal text alignment was applied successfully.");
        }
        else
        {
            System.out.println("Horizontal text alignment setup failed.");
        }

        //Setting up a custom vertical text alignment
        range.setVerticalAlignment(TextAlignment.TOP);

        //Verifing the vertical text alignment
        if (range.getVerticalAlignment().equals(TextAlignment.TOP))
        {
            System.out.println("A1:A3 range: a new vertical text alignment was applied successfully.");
        }
        else
        {
            System.out.println("A vertical text alignment setup failed.");
        }

        //Setting up custom text orientation
        range.setTextOrientation(TextOrientation.UPWARD);

        //Verifing text orientation
        if (range.getTextOrientation().equals(TextOrientation.UPWARD))
        {
            System.out.println("A1:A3 range: a new text orientation was applied successfully.");
        }
        else
        {
            System.out.println("A text orientation setup failed.");
        }

        application.close();
    }
}