import uno
import argparse

def force_update(doc_url, host, port):
    local_context = uno.getComponentContext()
    resolver = local_context.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_context)
    context = resolver.resolve(
        f"uno:socket,host={host},port={port};urp;StarOffice.ComponentContext")
    desktop = context.ServiceManager.createInstanceWithContext(
        "com.sun.star.frame.Desktop", context)
    
    properties = (
        uno.createUnoStruct("com.sun.star.beans.PropertyValue"),
        uno.createUnoStruct("com.sun.star.beans.PropertyValue")
    )
    properties[0].Name = "Hidden"
    properties[0].Value = True
    properties[1].Name = "MacroExecutionMode"
    properties[1].Value = 4
    
    doc = desktop.loadComponentFromURL(doc_url, "_blank", 0, properties)

    # Update embedded formulas
    embedded_objects = doc.EmbeddedObjects
    end_index = embedded_objects.Count - 1

    for index in range(end_index + 1):
        embedded_object = embedded_objects.getByIndex(index)
        model = embedded_object.Model
        if not model is None and model.supportsService("com.sun.star.formula.FormulaProperties"):
            # It is a formula object.
            xcoeo = embedded_object.ExtendedControlOverEmbeddedObject
            xcoeo.update()
    
    doc.store()
    # doc.close(True)

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("-d", "--doc_url", required=True)
    ap.add_argument("-H", "--host", required=True)
    ap.add_argument("-P", "--port", required=True)
    args = vars(ap.parse_args())
    force_update(args["doc_url"], args["host"], args["port"])
