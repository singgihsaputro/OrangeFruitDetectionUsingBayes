package range;

import com.jniwrapper.win32.jexcel.*;
import com.jniwrapper.win32.excel._Workbook;
import com.jniwrapper.win32.excel.Names;
import com.jniwrapper.win32.excel.Name;
import com.jniwrapper.win32.excel.Comment;
import com.jniwrapper.win32.automation.types.Variant;
import com.jniwrapper.win32.automation.types.BStr;
import com.jniwrapper.Int32;

public class GetRangeNamesExample
{
    public static void main(String[] args) throws Exception
    {
        Application application = new Application();
        final Workbook workbook = application.createWorkbook(null);

        final Worksheet activeWorksheet = workbook.getActiveWorksheet();
        workbook.getOleMessageLoop().doInvokeAndWait(new Runnable() {
            public void run() {
                Cell cell = activeWorksheet.getCell("A1");
                Comment comment = cell.getPeer().getComment();
                System.out.println("comment = " + comment);
            }
        });

        application.getOleMessageLoop().doInvokeAndWait(new Runnable() {
            public void run() {
                _Workbook workbookPeer = workbook.getNativePeer();

                Names names = workbookPeer.getNames();
                Int32 numberOfNames = names.getCount();
                System.out.println("numberOfNames = " + numberOfNames);

                Variant unspecified = Variant.createUnspecifiedParameter();
                for (int i = 1; i <= numberOfNames.getValue(); i++) {
                    Name name = names.item(new Variant(i), unspecified, unspecified);
                    BStr nameValue = name.getValue();
                    System.out.println("nameValue = " + nameValue);

                    BStr nameLocal = name.getNameLocal();
                    System.out.println("nameLocal = " + nameLocal);
                }
            }
        });

        application.close();
    }
}
