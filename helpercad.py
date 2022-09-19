import win32com.client
import pythoncom

acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument
acadModel = doc.ModelSpace

def APoint(x, y, z = 0):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))

def aDouble(xyz):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (xyz))

def aVariant(vObject):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, (vObject))

doc.Utility.Prompt("Hello from Python\n")

ss_name = 'SS999'
try:
    doc.SelectionSets.Item(ss_name).Delete()
except:
    print("Delete selection failed")

ss1 = doc.SelectionSets.Add(ss_name)
ss1.SelectOnScreen()

for i in range(int(ss1.Count)):
    print(ss1.Item(i).ObjectName)
    entity = ss1.Item(i)
    name = entity.EntityName
    if name == 'AcDbBlockReference':
        HasAttributes = entity.HasAttributes
        if HasAttributes:
            for attrib in entity.GetAttributes():
                print("  {}: {}".format(attrib.TagString, attrib.TextString))



