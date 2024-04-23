from uno import getComponentContext, createUnoStruct
import argparse

def open_document(doc_url, host, port):
    local_context = getComponentContext()
    resolver = local_context.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_context)
    context = resolver.resolve(
        f"uno:socket,host={host},port={port};urp;StarOffice.ComponentContext")
    desktop = context.ServiceManager.createInstanceWithContext(
        "com.sun.star.frame.Desktop", context)
    
    properties = (
        createUnoStruct("com.sun.star.beans.PropertyValue"),
        createUnoStruct("com.sun.star.beans.PropertyValue")
    )
    properties[0].Name = "Hidden"
    properties[0].Value = True
    properties[1].Name = "MacroExecutionMode"
    properties[1].Value = 4
    
    doc = desktop.loadComponentFromURL(doc_url, "_blank", 0, properties)

    # Update document indexes
    if doc.supportsService("com.sun.star.text.GenericTextDocument"):
        indexes = doc.getDocumentIndexes()
        for i in range(indexes.getCount()):
            indexes.getByIndex(i).update()

    doc.store()
    # doc.close(True)


if __name__ == '__main__':
	# construct the argument parse and parse the arguments
    ap = argparse.ArgumentParser()
    ap.add_argument("-d", "--doc_url", required=True)
    ap.add_argument("-H", "--host", required=True)
    ap.add_argument("-P", "--port", required=True)
    args = vars(ap.parse_args())
    open_document(args["doc_url"], args["host"], args["port"])
