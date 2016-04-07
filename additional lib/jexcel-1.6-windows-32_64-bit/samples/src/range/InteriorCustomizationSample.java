/*
 * Copyright (c) 2000-2009 TeamDev Ltd. All rights reserved.
 * TeamDev PROPRIETARY and CONFIDENTIAL.
 * Use is subject to license terms.
 */
package range;

import com.jniwrapper.win32.jexcel.*;
import com.jniwrapper.win32.jexcel.format.InteriorPattern;

import java.awt.*;

/**
 * This sample demonstrates how to obtain and change interior settings.
 *
 * @author Vladimir Kondrashchenko
 */
public class InteriorCustomizationSample
{
    public static void main(String[] args) throws ExcelException
    {
        Application application = new Application();
        GenericWorkbook workbook = application.createWorkbook(null);
        Worksheet worksheet = workbook.getWorksheet(1);

        Range range = worksheet.getRange("A1:D12");

        //Getting default interior setting
        Color interiorColor = range.getInteriorColor();
        InteriorPattern interiorPattern = range.getInteriorPattern();
        Color interiorPatternColor = range.getInteriorPatternColor();

        //Printing default interior settings
        System.out.println("Default interior settings:");
        System.out.println("Interior color: " + interiorColor);
        if (interiorPattern.equals(InteriorPattern.NONE))
        {
            System.out.println("Interior pattern is not set up.");
        }
        else
        {
            System.out.println("Interior pattern code is: " + interiorPattern.getLongValue());
        }
        System.out.println("Interior pattern color: " + interiorPatternColor);

        //Changing interior settings
        range.setInteriorColor(Color.BLUE);
        range.setInteriorPattern(InteriorPattern.DOWN);
        range.setInteriorPatternColor(Color.RED);

        //Getting new interior settings
        interiorColor = range.getInteriorColor();
        interiorPattern = range.getInteriorPattern();
        interiorPatternColor = range.getInteriorPatternColor();

        //Printing new interior settings
        System.out.println("\nNew interior settings:");
        System.out.println("Interior color: " + interiorColor);
        if (interiorPattern.equals(InteriorPattern.DOWN))
        {
            System.out.println("\"DOWN\" interior pattern is set up.");
        }
        else
        {
            System.out.println("Interior pattern code is: " + interiorPattern.getLongValue());
        }
        System.out.println("Interior pattern color: " + interiorPatternColor);

        application.close();
    }
}