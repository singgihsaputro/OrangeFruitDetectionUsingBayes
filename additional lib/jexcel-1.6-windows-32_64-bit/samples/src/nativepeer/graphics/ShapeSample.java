/*
 * Copyright (c) 2000-2009 TeamDev Ltd. All rights reserved.
 * TeamDev PROPRIETARY and CONFIDENTIAL.
 * Use is subject to license terms.
 */
package nativepeer.graphics;

import com.jniwrapper.Int32;
import com.jniwrapper.win32.automation.types.BStr;
import com.jniwrapper.win32.automation.types.Variant;
import com.jniwrapper.win32.com.IUnknown;
import com.jniwrapper.win32.excel.*;
import com.jniwrapper.win32.excel.impl.ChartObjectImpl;
import com.jniwrapper.win32.excel.impl.ChartObjectsImpl;
import com.jniwrapper.win32.excel.impl._ChartImpl;
import com.jniwrapper.win32.jexcel.Application;
import com.jniwrapper.win32.jexcel.Workbook;

public class ShapeSample
{

    private static Workbook wb;

    public static void main(String[] args) throws Exception
    {
        Application app = new Application();
        wb = app.createWorkbook("Test");
        wb.getOleMessageLoop().doInvokeAndWait(new Runnable()
        {
            public void run()
            {
                performActions();
            }
        });

        app.close();
        System.out.println("Done");
    }

    public static void performActions()
    {
        // NOTE: user executing export function must be interactive to be able
        // use graphic filters from registry
        performSheetTasks("Sheet1");
        performAppTasks();
    }

    public static void performSheetTasks(String sheetName)
    {
        _Worksheet ws = wb.getWorksheet(sheetName).getPeer();
        IUnknown disp = ws.chartObjects(Variant.createUnspecifiedParameter(),
                new Int32(0));
        ChartObjects objects = new ChartObjectsImpl(disp);
        for (int i = 1; i <= objects.getCount().getValue(); i++)
        {
            String fileName = wb.getFile().getParent() + "\\"
                    + wb.getFile().getName() + "_" + sheetName + "_" + i
                    + ".jpg";
            IUnknown chartUnknown = ws.chartObjects(new Variant(i),
                    new Int32(0));
            ChartObject chartObject = new ChartObjectImpl(chartUnknown);
            _Chart chart = chartObject.getChart();
            saveChartAsJPEG(chart, fileName);
        }
    }

    protected static void performAppTasks()
    {
        _Workbook _wb = wb.getNativePeer();
        Sheets sheets = _wb.getCharts();
        for (int i = 1; i <= sheets.getCount().getValue(); i++)
        {
            IUnknown sheet = sheets.getItem(new Variant(i));
            _Chart chart = new _ChartImpl(sheet);
            saveChartAsJPEG(chart, wb.getFile().getParent() + "\\"
                    + wb.getFile().getName() + "_" + chart.getName().getValue()
                    + "_" + i + ".jpg");
        }
    }

    protected static void saveChartAsJPEG(_Chart chart, String location)
    {
        chart.export(new BStr(location), new Variant("JPEG"), Variant
                .createUnspecifiedParameter());
    }
}