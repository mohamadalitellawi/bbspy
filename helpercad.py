
from pathlib import Path
import win32com.client
import pythoncom

# acad = win32com.client.Dispatch("AutoCAD.Application")
# doc = acad.ActiveDocument
# acadModel = doc.ModelSpace

def APoint(x, y, z = 0):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))

def aDouble(xyz):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (xyz))

def aVariant(vObject):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, (vObject))





def main():
    active_document = get_cad_active_doc()
#    print(link_barinfo_to_block(active_document))
#    print(get_block_attributes(active_document))
#    print(get_dynamic_block_data(active_document))
    export_image(active_document)
    active_document= None



def export_image(doc):
    ss1 = get_selection_from_screen(doc)
    for obj in ss1:
        p1 = p2 = []
        [p1,p2] = obj.GetBoundingBox(p1,p2)
        p1 = aDouble(p1)
        p2 = aDouble(p2)
        print(p1,p2)
        doc.Application.ZoomWindow(p1,p2)
        filename = "Z:\\03 - Design\\ 300-ALI\\_From Omar\\16-10-2022\\" + "shape" + obj.Handle
        doc.export(filename, "WMF", ss1)
        doc.Application.ZoomPrevious()


def link_barinfo_to_block(doc):
    returnObj = None
    basePnt = None
    [returnObj, basePnt] = doc.Utility.GetEntity (returnObj, basePnt, "Select an object")
    return returnObj.Name


def get_dynamic_block_data(doc):
    ss1 = get_selection_from_screen(doc)
    if ss1 is None: return
    result = []
    for i in range(int(ss1.Count)):
        entity = ss1.Item(i)
        type = entity.ObjectName
        id = entity.ObjectID
        handle = entity.Handle
        if type == 'AcDbBlockReference':
            if not entity.IsDynamicBlock:
                continue
            block_info = {}
            block_info['type'] = type
            block_info['id'] = id
            block_info['handle'] = handle
            block_info['block_name'] = entity.Name
            block_info['block_effective_name'] = entity.EffectiveName
            block_info['properties'] = []

            block_properties = entity.GetDynamicBlockProperties()
            for block_property in block_properties:
                block_info['properties'].append({block_property.PropertyName : block_property.Value})

            result.append(block_info)
    if len(result) > 0:
        return result

def get_block_attributes(doc):

    ss1 = get_selection_from_screen(doc)
    if ss1 is None: return
    result = []
    for i in range(int(ss1.Count)):
#        print(ss1.Item(i).ObjectName)
        entity = ss1.Item(i)
        type = entity.ObjectName
        id = entity.ObjectID
        handle = entity.Handle
        if type == 'AcDbBlockReference':
            HasAttributes = entity.HasAttributes
            if HasAttributes:
                bar_info = {}
                bar_info['type'] = type
                bar_info['id'] = id
                bar_info['handle'] = handle
                bar_info['block_name'] = entity.Name
                bar_info['block_effective_name'] = entity.EffectiveName

                bar_info['attributes'] = []
                
                for attrib in entity.GetAttributes():
                    #print("  {}: {}".format(attrib.TagString, attrib.TextString))
                    bar_info['attributes'].append({attrib.TagString:attrib.TextString})
                result.append(bar_info)
    if len(result) > 0:
        return result


def get_cad_active_doc():
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application")
        doc = acad.ActiveDocument
        doc.Utility.Prompt("Hello from Python\n")
        print(doc.Name)
        return doc
    except:
        print('Error!:\t', 'we have a problem to connect with autocad')

def get_selection_from_screen(doc):
    ss_name = 'SS999'
    try:
        doc.SelectionSets.Item(ss_name).Delete()
    except:
        print("Delete selection failed")

    ss1 = doc.SelectionSets.Add(ss_name)
    try:
        ss1.SelectOnScreen()
        if int(ss1.Count) > 0:
            return ss1
    except:
        print('Error!:\t', 'we have a problem to select from autocad')

if __name__ == '__main__':
    main()