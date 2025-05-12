# json-to-odt

Generates Openoffice odt from json generated through **gsheet-to-json**

copy ```conf/config.yml.dist``` as ```conf/config.yml``` and do not commit the copied file

## software and tools required

1.  Python 3.8 or higher - <https://www.python.org/downloads/>
2.  Git -  <https://git-scm.com/downloads>
3.  LibreOffice - <https://www.libreoffice.org/download/download/>. You can work without LibreOffice installed on Linux and Windows both, but you will not be able to generate indexes and pdf from the generated odt

## we need an application macro for odt index generation to work
Create the following two macros from *Tools->Macros->Edit Macros...*

* Container **My Macros & Dialogs**
* Application **Standard**
* Module **Module1**
* Function **open_document(String)**

```
Sub open_document(docUrl As String)

	Dim doc As Object
	Dim properties(1) As New com.sun.star.beans.PropertyValue

	properties(0).Name = "Hidden"
	properties(0).value  = true
	properties(1).Name = "MacroExecutionMode"
	properties(1).value  = 4

	doc = StarDesktop.loadComponentFromURL(docUrl, "_blank", 0, properties)

	'''Update indexes, such as for the table of contents'''
	Dim i As Integer

    With doc
        If .supportsService("com.sun.star.text.GenericTextDocument") Then
            For i = 0 To .getDocumentIndexes().count - 1
                .getDocumentIndexes().getByIndex(i).update()
            Next i
        End If
    End With

	doc.store()
	doc.close(True)

End Sub
```



* Container **My Macros & Dialogs**
* Application **Standard**
* Module **Module1**
* Function **force_update(String)**

```
Sub force_update(docUrl As String)
	
	Dim doc As Object
	Dim properties(1) As New com.sun.star.beans.PropertyValue

	properties(0).Name = "Hidden"
	properties(0).value  = true
	properties(1).Name = "MacroExecutionMode"
	properties(1).value  = 4

	doc = StarDesktop.loadComponentFromURL(docUrl, "_blank", 0, properties)
	'''doc = ThisComponent

	'''Update embedded formulas'''
	Dim oEmbeddedObjects as Variant: oEmbeddedObjects = doc.EmbeddedObjects
	Dim nIndex as Long
	Dim nEndIndex as Long: nEndIndex = oEmbeddedObjects.Count-1

	rem like green handle status
	Dim oEmbeddedObject as Variant:
	rem like edit status 
	Dim oModel as Variant: 
	rem oExtendedControlOverEmbeddedObject
	Dim oXCOEO as Variant: 

	For nIndex=0 to nEndIndex
		oEmbeddedObject = oEmbeddedObjects.getByIndex(nIndex)
		oModel = oEmbeddedObject.Model: rem might be empty; css.comp.math.FormulaDocument
		If Not(isEmpty(oModel)) then
			If oModel.supportsService("com.sun.star.formula.FormulaProperties") then
				rem It is a formula object.
				oXCOEO = oEmbeddedObject.ExtendedControlOverEmbeddedObject
				oXCOEO.update()
			End If	
		End If
	Next nIndex
	
	doc.store()
	doc.close(True)
	
End Sub
```


## macro security
For now, we need to allow macros in LibreOffice
* From **Tools->Options menu**
* Go to **LibreOffice->Security**
* Go to **Macro Security**
* Select **Low**

## Useful links
some useful links for json-to-odt
* https://mashupguide.net/1.0/html/ch17s04.xhtml
* https://github.com/eea/odfpy/tree/master/examples
* https://glot.io/snippets/f0nuzv7b8k


## Linux note:
You may get a warning like this
> Warning: failed to launch javaldx - java may not function correctly

To resolve this run

```sudo apt-get install libreoffice-java-common default-jre```

## Linux usage:
you can run the bash script *odt-from-gsheet.sh* in the root folder. It takes only one parameter - the name of the gsheet

```./odt-from-gsheet.sh name-of-the-gsheet```

or you can run the python script this way (NOT PREFERRED)
```
DOCULABORATION_BASE="/home/asif/projects/asif@github/doculaboration"
cd ${DOCULABORATION_BASE}
cd ./json-to-odt/src
./odt-from-json.py --config '../conf/config.yml'
```

## Windows usage:
you can run the windows command script *odt-from-gsheet.bat* in the root folder. It takes only one parameter - the name of the gsheet

```odt-from-gsheet.bat name-of-the-gsheet```

or you can run the python script this way (NOT PREFERRED)
```
set DOCULABORATION_BASE="D:\projects\asif@github\doculaboration"
cd %DOCULABORATION_BASE%
cd ./json-to-odt/src
python odt-from-json.py --config "../conf/config.yml"
```
