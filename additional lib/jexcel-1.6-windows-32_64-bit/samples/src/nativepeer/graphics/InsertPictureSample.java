/*
 * Copyright (c) 2000-2009 TeamDev Ltd. All rights reserved.
 * TeamDev PROPRIETARY and CONFIDENTIAL.
 * Use is subject to license terms.
 */
package nativepeer.graphics;

import com.jniwrapper.Int32;
import com.jniwrapper.win32.automation.IDispatch;
import com.jniwrapper.win32.automation.types.BStr;
import com.jniwrapper.win32.automation.types.Variant;
import com.jniwrapper.win32.excel.Pictures;
import com.jniwrapper.win32.excel._Worksheet;
import com.jniwrapper.win32.excel.impl.PicturesImpl;
import com.jniwrapper.win32.jexcel.Application;
import com.jniwrapper.win32.jexcel.GenericWorkbook;
import com.jniwrapper.win32.jexcel.Worksheet;

public class InsertPictureSample
{
    public static void main(String[] args) throws Exception
    {

        Application application = new Application();
        application.setVisible(true);
        GenericWorkbook workbook = application.createWorkbook("Picture");

        final Worksheet activeWorksheet = workbook.getActiveWorksheet();
        application.getOleMessageLoop().doInvokeAndWait(new Runnable()
        {
            public void run()
            {
                _Worksheet worksheet = activeWorksheet.getPeer();
                IDispatch picturesDispatch = worksheet.pictures(Variant.createUnspecifiedParameter(), new Int32(0));
                Pictures pictures = new PicturesImpl(picturesDispatch);
                pictures.insert(new BStr("C:\\picture.png"), Variant.createUnspecifiedParameter());
            }
        });

        System.out.println("Press 'Enter' to terminate application.");
        System.in.read();
        application.close();
    }
}