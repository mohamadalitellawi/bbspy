

from pathlib import Path
import win32com.client
import pythoncom
from model_bar_block import BarBlockData
from model_bar_info import BarInfoBlock

# acad = win32com.client.Dispatch("AutoCAD.Application")
# doc = acad.ActiveDocument
# acadModel = doc.ModelSpace

def APoint(x, y, z = 0):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))

def aDouble(xyz):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (xyz))

def aVariant(vObject):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, (vObject))


BAR_INFO_HEADER = ['BAR_MARK', 'BAR_INFO','BAR_LOCATION','BAR_SHAPE','VARIABLES']


def main():
    active_document = get_cad_active_doc()
#    print(link_barinfo_to_block(active_document))
#    print(get_block_attributes(active_document))
#    print(get_dynamic_block_data(active_document))
#    export_image(active_document)
#    get_point()
#    print(get_select_all(active_document))
    rename_all_dyn_blocks(active_document)
    active_document= None

def rename_all_dyn_blocks(doc,suffix = ''):
    '''
    ss1 = get_select_all(doc)
    if ss1 is None: return
    suffix = 'XX' + str(suffix)
    for i in range(int(ss1.Count)):
        entity = ss1.Item(i)
        type = entity.ObjectName
        id = entity.ObjectID
        handle = entity.Handle
        if type == 'AcDbBlockReference':
            if entity.IsDynamicBlock:
                block_new_name = entity.EffectiveName + suffix
                entity.Name = 'tttttt'
    '''
    suffix = 'XX' + str(suffix)
    for block in doc.Blocks:
        entity = block
        type = entity.ObjectName
        #print(type)
        if type == 'AcDbBlockTableRecord':
            if entity.IsDynamicBlock:
                block_new_name = entity.Name + suffix
                entity.Name = block_new_name
    doc.Utility.Prompt ("\nDone!")


def get_point():
    doc = get_cad_active_doc()
    returnObj = None
    basePnt = aDouble((0,0,0))
    try:
        returnObj = doc.Utility.GetPoint (basePnt, "select point")
        print(returnObj)
    except:
        print("failed to get point from drawing")
        raise


def export_image(doc, filename):
    ss1 = get_selection_from_screen(doc)
    for obj in ss1:
        p1 = p2 = []
        [p1,p2] = obj.GetBoundingBox(p1,p2)
        p1 = aDouble(p1)
        p2 = aDouble(p2)
        print(p1,p2)
        doc.Application.ZoomWindow(p1,p2)
        doc.export(filename, "WMF", ss1)
        doc.Application.ZoomPrevious()


def link_barinfo_to_block(doc):
    returnObj = None
    basePnt = None
    [returnObj, basePnt] = doc.Utility.GetEntity (returnObj, basePnt, "Select an object")
    return returnObj.EffectiveName


def get_cad_entity(doc, message = "Select an object: "):
    returnObj = None
    basePnt = None
    try:
        [returnObj, basePnt] = doc.Utility.GetEntity (returnObj, basePnt, message)
        return returnObj
    except:
        print("failed to select cad entity")


def get_bar_block_data(entity):
    type = entity.ObjectName
    id = entity.ObjectID
    handle = entity.Handle
    if type == 'AcDbBlockReference':
        if not entity.IsDynamicBlock:
            return
        bar_block = BarBlockData()
        bar_block.name = entity.EffectiveName
        bar_block.id = id
        bar_block.handle = handle
        bar_block.insertion_point = entity.InsertionPoint
        block_properties = entity.GetDynamicBlockProperties()
        for block_property in block_properties:
            bar_block.dimensions.append({block_property.PropertyName : block_property.Value})
        return bar_block


def add_layer(doc, layername):
    try:
        check_if_exist = False
        for layer in doc.Layers:
            if layername == layer.Name:
                check_if_exist = True
        if not check_if_exist:
            doc.Layers.Add(layername)
    except:
        print("failed to add layer in drawing")

def insert_block(doc, shape_blockname,bar_data:BarInfoBlock,scale = 1.0, message = 'Select Block Insertion Point: '):
    doc.StartUndoMark()
    returnObj = None
    basePnt = aDouble(bar_data.insertion_point)
    try:
        returnObj = doc.Utility.GetPoint (basePnt, message)
        #print(returnObj)
        insertionPnt = aDouble(returnObj)
        blockRefObj = doc.ModelSpace.InsertBlock(insertionPnt, shape_blockname, scale, scale, scale, 0)
        doc.Utility.Prompt ("Shape Inserted!")
        block_attributes = blockRefObj.GetAttributes()
        for attrib in block_attributes:
            for k,v in bar_data.attributes.items():
                if attrib.TagString == k:
                    attrib.TextString = v
                    break
    except:
        print("failed to get point from drawing")
        doc.Utility.Prompt ("Error Happend - Inserting Shape Block")
        blockRefObj.Delete()
    finally:
        doc.EndUndoMark()

def draw_circles(doc, layername, centers,radius):
    for point in centers:
        centerPoint = aDouble(point)
        circleObj = doc.ModelSpace.AddCircle(centerPoint, radius)
        circleObj.Layer = layername


def get_barinfo_list(doc):
    ss1 = get_selection_from_screen(doc)
    if ss1 is None: return
    result = []
    for i in range(int(ss1.Count)):
        entity = ss1.Item(i)
        type = entity.ObjectName
        if type != 'AcDbBlockReference':
            continue
        if not entity.HasAttributes:
            continue
        if entity.EffectiveName != 'BAR-INFO':
            continue

        bar_info = BarInfoBlock()
        bar_info.id = entity.ObjectID
        bar_info.handle = entity.Handle
        bar_info.name = entity.EffectiveName
        bar_info.insertion_point = entity.InsertionPoint

        for attrib in entity.GetAttributes():
            bar_info.attributes[attrib.TagString] = attrib.TextString

        result.append(bar_info)

    if len(result) > 0:
        print(f'total selected {len(result)} bar info blocks')
        return result

def update_bar_info(entity, bar_block:BarBlockData, bar_parameters):
    type = entity.ObjectName
    id = entity.ObjectID
    handle = entity.Handle
    bar_info = BarInfoBlock()
    if type == 'AcDbBlockReference':
        if not entity.HasAttributes:
            return
        bar_info.name = entity.EffectiveName
        bar_info.id = id
        bar_info.handle = handle
        bar_info.attributes['BAR_SHAPE'] = bar_block.get_original_name()
        bar_info.insertion_point = entity.InsertionPoint
        parameter = distance = ''
        for attrib in entity.GetAttributes():
            if attrib.TagString == 'BAR_SHAPE':
                attrib.TextString = bar_block.get_original_name()
            if attrib.TagString not in BAR_INFO_HEADER:
                attrib.TextString = '0'
            for x in bar_parameters:
                if x['LEGNTH'] == attrib.TagString:
                    parameter = x['PARAMETER']
                    distance = x['REMARK']
                    attr_field = fr'%<\AcObjProp Object(%<\_ObjId {bar_block.id}>%).Parameter({parameter}).{distance} \f "%lu2%pr0">%'
                    #print(attr_field)
                    attrib.TextString = attr_field
                    bar_info.attributes[str(attrib.TagString)] = attr_field
    return bar_info


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

def get_select_all(doc):
    ss_name = 'SS999'
    acSelectionSetAll = 5

    '''
acSelectionSetWindow=0
acSelectionSetCrossing=1
acSelectionSetPrevious=3
acSelectionSetLast=4
acSelectionSetAll=5
'''

    try:
        doc.SelectionSets.Item(ss_name).Delete()
    except:
        print("Delete selection failed")

    ss1 = doc.SelectionSets.Add(ss_name)
    try:
        ss1.Select(acSelectionSetAll)
        if int(ss1.Count) > 0:
            return ss1
    except:
        print('Error!:\t', 'we have a problem to select from autocad')

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