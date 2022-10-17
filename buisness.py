from pathlib import Path
import helpercad
import helperfile
from model_bar_block import BarBlockData
from model_bar_info import BarInfoBlock

FILES = {
    'source_folder' : Path('C:\\BBS_SOURCE'),
    'images_folder' : Path('C:\\BBS_SOURCE\\IMG'),
    'parameters': Path('C:\\BBS_SOURCE\\PARAMETER NUMBERING.CSV')
}

SHAPE_SCALE_FACTOR = 1.0

parameters = helperfile.get_parameters(FILES['parameters'])

def main():
    link_Bar_Info()


def link_Bar_Info():
    doc = helpercad.get_cad_active_doc()
    if doc is None:
        return
    try:
        while (True):
            bar_block = get_bar_block(doc)
            bar_parameters = [x for x in parameters if x['BLOCK NAME'] == bar_block.get_original_name()]
            bar_info = update_bar_info(doc,bar_block,bar_parameters)
            helpercad.insert_block(doc,bar_block.get_bar_shapename(),bar_info,SHAPE_SCALE_FACTOR, 'Select Shape Insertion Point: ')
    finally:
        doc = None


def get_bar_block(doc):
    entity = helpercad.get_cad_entity(doc, "Select Bar Block: ")
    bar_block = helpercad.get_bar_block_data(entity)
    return bar_block

def update_bar_info(doc, bar_block:BarBlockData,bar_parameters):
    entity = helpercad.get_cad_entity(doc, "Select Bar Info Block: ")
    bar_info = helpercad.update_bar_info(entity,bar_block,bar_parameters)
    return bar_info


def get_shape_insertion_point():
    pass



if __name__ == '__main__':
    main()