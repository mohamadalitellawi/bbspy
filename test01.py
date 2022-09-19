import win32com.client
import pythoncom

acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument
acadModel = doc.ModelSpace

#https://thgeomacademy.wordpress.com/python-autocad-selectionsets/
#https://www.supplychaindataanalytics.com/selectionset-object-in-autocad-with-python/


def APoint(x, y, z = 0):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))

def aDouble(xyz):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (xyz))

def aVariant(vObject):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, (vObject))

l1 = acadModel.AddLine(APoint(0, 0, 0),APoint(1000, 1000, 0))
l2 = acadModel.AddLine(APoint(1000, 1000, 0),APoint(2000, 0, 0))

ss_name = 'SS2'
try:
    doc.SelectionSets.Item(ss_name).Delete()
except:
    print("Delete selection failed")

ss1 = doc.SelectionSets.Add(ss_name)

ss1.AddItems(aVariant([l1, l2]))
print(ss1.Name)
print(ss1.Count)
print(ss1.Item(0).ObjectName)
ss1.SelectOnScreen()

for i in range(int(ss1.Count)):
    print(ss1.Item(i).ObjectName)

