/*
 * Copyright (c) 2000-2009 TeamDev Ltd. All rights reserved.
 * TeamDev PROPRIETARY and CONFIDENTIAL.
 * Use is subject to license terms.
 */
package range;

import com.jniwrapper.win32.jexcel.*;
import com.jniwrapper.win32.jexcel.format.Border;

import java.awt.*;

/**
 * This sample demonstrates how to obtain and change border setting.
 *
 * @author Vladimir Kondrashchenko
 */
public class BorderCustomizationSample
{
    public static void main(String[] args) throws ExcelException
    {
        Application application = new Application();
        GenericWorkbook workbook = application.createWorkbook(null);
        Worksheet worksheet = workbook.getWorksheet(1);

        Range range = worksheet.getRange("A1:D12");

        //Getting top border
        Border topBorder = range.getBorder(Border.Kind.EDGETOP);

        //Getting border settings
        Color borderColor = topBorder.getColor();
        Border.LineStyle lineStyle = topBorder.getLineStyle();
        Border.LineWeight lineWeight = topBorder.getWeight();

        //Printing border settings
        System.out.println("Top border settings:");
        System.out.println("Color: " + borderColor);
        System.out.println("Line style: " + lineStyle);
        System.out.println("Line weight: " + lineWeight);

        //Setting new border style
        Border border = new Border();
        border.setColor(Color.CYAN);
        border.setLineStyle(Border.LineStyle.DASHDOT);
        border.setWeight(Border.LineWeight.MEDIUM);

        range.setBorder(Border.Kind.EDGETOP, border);
        range.setBorder(Border.Kind.EDGELEFT, border);

        application.close();
    }
}