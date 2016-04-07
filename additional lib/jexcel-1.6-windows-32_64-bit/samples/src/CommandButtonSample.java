import com.jniwrapper.Int32;
import com.jniwrapper.win32.automation.Automation;
import com.jniwrapper.win32.automation.IDispatch;
import com.jniwrapper.win32.automation.types.BStr;
import com.jniwrapper.win32.automation.types.Variant;
import com.jniwrapper.win32.excel.Shape;
import com.jniwrapper.win32.excel.Shapes;
import com.jniwrapper.win32.jexcel.Worksheet;
import com.jniwrapper.win32.jexcel.ui.JWorkbook;

import javax.swing.*;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;

/**
 * Sample demonstrates how to add a custom command button to a worksheet.
 */
public class CommandButtonSample
{
    public static void main(String[] args) throws Exception
    {
        final JWorkbook workbook = new JWorkbook();
        workbook.addWorksheet("test");

        final Worksheet worksheet = workbook.getWorksheet("Test");

        JFrame frame = new JFrame();
        frame.setContentPane(workbook);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(800, 600);
        frame.addWindowListener(new WindowAdapter()
        {
            public void windowClosing(WindowEvent e)
            {
                workbook.close();
            }
        });

        worksheet.getApplication().getOleMessageLoop().doInvokeAndWait(new Runnable()
        {
            public void run()
            {
                Shapes shapes = worksheet.getPeer().getShapes();

                Variant unspecifiedParameter = Variant.createUnspecifiedParameter();
                Shape testButton = shapes.addOLEObject(new Variant("Forms.CommandButton.1"),
                        unspecifiedParameter,
                        unspecifiedParameter,
                        unspecifiedParameter,
                        unspecifiedParameter,
                        unspecifiedParameter,
                        unspecifiedParameter,
                        new Variant(140),
                        new Variant(18),
                        new Variant(100),
                        new Variant(25));

                shapes.setAutoDelete(false);
                shapes.release();

                try
                {
                    testButton.setName(new BStr("Test"));
                    IDispatch obj = worksheet.getPeer().OLEObjects(new Variant("Test"), new Int32(0));

                    Automation objAutomation = new Automation(obj, true);
                    IDispatch intObj = objAutomation.getProperty("Object").getPdispVal();
                    if (!intObj.isNull())
                    {
                        Automation intObjAutomation = new Automation(intObj, true);
                        intObjAutomation.setProperty("Caption", "Click Me!");
                        intObjAutomation.release();
                    }
                    objAutomation.release();

                    intObj.setAutoDelete(false);
                    intObj.release();

                    obj.setAutoDelete(false);
                    obj.release();
                }
                catch (Exception ex)
                {
                    ex.printStackTrace();
                }

                System.out.println("Command button name: " + testButton.getName());
                testButton.setAutoDelete(false);
                testButton.release();
            }
        });

        frame.setVisible(true);
    }
}
