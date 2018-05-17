### **Word转PDF**
---

 > 使用jacob操作offic。

---

要求:

- [x] windowOS
- [x] officeWord 2007 +
- [x] SaveAsPDFandXPS插件

---

步骤:
1. 将.dll文件放入jre/bin下
2. 安装SaveAsPDFandXPS.exe
3. 导入jacob.jar
4. 引入代码
```
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

import java.io.File;

public class Word2PDF {


    private static final int wdFormatPDF = 17;// PDF 格式
    /**
     * @param sourceFile
     * @param PDFPath
     */
    public static void wordToPDF(String sourceFile , String PDFPath){

        ActiveXComponent app = null;
        Dispatch doc = null;
        try {
            app = new ActiveXComponent("Word.Application");
            app.setProperty("Visible", new Variant(false));
            Dispatch docs = app.getProperty("Documents").toDispatch();

            doc = Dispatch.call(docs,  "Open" , sourceFile).toDispatch();
            File tofile = new File(PDFPath);
            if (tofile.exists()) {
                tofile.delete();
            }
            Dispatch.call(doc,"SaveAs", PDFPath, wdFormatPDF);
        } catch (Exception e) {
            System.out.println(e.getMessage());
        } finally {
            Dispatch.call(doc,"Close",false);
            if (app != null)
                app.invoke("Quit", new Variant[] {});
        }

        ComThread.Release();
    }

}
```
