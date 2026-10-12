"""Run with LibreOffice's python: update fields/indexes, save .docx and .pdf.
usage: python lo_finalize.py in.docx out.docx out.pdf"""
import sys, time, subprocess, uno
from com.sun.star.beans import PropertyValue


def pv(n, v):
    p = PropertyValue(); p.Name = n; p.Value = v; return p


src, out_docx, out_pdf = sys.argv[1:4]
proc = subprocess.Popen(["/Applications/LibreOffice.app/Contents/MacOS/soffice", "--headless", "--invisible", "--norestore",
                         "--accept=socket,host=127.0.0.1,port=2083;urp;"])
ctx = None
for _ in range(60):
    try:
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
        ctx = resolver.resolve("uno:socket,host=127.0.0.1,port=2083;urp;StarOffice.ComponentContext")
        break
    except Exception:
        time.sleep(1)
desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(src), "_blank", 0, (pv("Hidden", True),))
doc.getTextFields().refresh()
idx = doc.getDocumentIndexes()
for _ in range(2):  # twice: page numbers settle after the indexes themselves take space
    for i in range(idx.getCount()):
        idx.getByIndex(i).update()
    doc.getTextFields().refresh()
print("indexes:", idx.getCount())
doc.storeToURL(uno.systemPathToFileUrl(out_docx), (pv("FilterName", "MS Word 2007 XML"),))
doc.storeToURL(uno.systemPathToFileUrl(out_pdf), (pv("FilterName", "writer_pdf_Export"),))
doc.close(True)
try:
    desktop.terminate()
except Exception:
    pass
proc.wait(timeout=30)
