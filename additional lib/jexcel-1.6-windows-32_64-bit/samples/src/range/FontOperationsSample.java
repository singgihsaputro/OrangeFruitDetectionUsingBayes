/*
 * Copyright (c) 2000-2009 TeamDev Ltd. All rights reserved.
 * TeamDev PROPRIETARY and CONFIDENTIAL.
 * Use is subject to license terms.
 */
package range;

import com.jniwrapper.win32.jexcel.*;
import com.jniwrapper.win32.jexcel.format.Font;

import java.awt.*;

/**
 * This sample demonstrates how to customize text font.
 *
 * @author Vladimir Kondrashchenko
 */
public class FontOperationsSample
{
    public static void main(String[] args) throws ExcelException
    {
        Application application = new Application();
        GenericWorkbook workbook = application.createWorkbook(null);
        Worksheet worksheet = workbook.getWorksheet(1);

        Range range = worksheet.getRange("A1:D12");

        Font font = range.getFont();

        //Printing default font settings
        System.out.println("Default font settings:");
        printFontAttributes(font);

        //Changing font name
        font.setName("Courier New");

        //Changing font styles
        font.setBold(true);
        font.setStrikethrough(true);
        font.setUnderlineStyle(Font.UnderlineStyle.DOUBLE);

        //Changing font color
        font.setColor(Color.ORANGE);

        //Applying new font setting
        range.setFont(font);

        //Printing new font settings
        System.out.println("\nNew font settings:");
        printFontAttributes(range.getFont());

        application.close();
    }

    public static void printFontAttributes(Font font)
    {
        System.out.println("Font name: " + font.getName());
        System.out.println("Font size: " + font.getSize());
        System.out.println("Font styles: ");

        if (font.isBold())
        {
            System.out.println("\tBold");
        }
        else
        {
            System.out.println("\tNot bold");
        }

        if (font.isItalic())
        {
            System.out.println("\tItalic");
        }
        else
        {
            System.out.println("\tNot Italic");
        }

        if (font.isStrikethrough())
        {
            System.out.println("\tStriked through");
        }
        else
        {
            System.out.println("\tNot striked through");
        }

        System.out.println("Font underline style: " + font.getUnderlineStyle());

        System.out.println("Font alignment: " + font.getAlignment());

        System.out.println("Font color: " + font.getColor());
    }
}