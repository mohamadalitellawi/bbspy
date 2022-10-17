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

ERROR_LAYER_NAME = '_BBS_ERROR'
ERROR_CIRCLE_RADIUS = 500

parameters = helperfile.get_parameters(FILES['parameters'])

def main():
    #link_Bar_Info()
    check_bbs()

def check_bbs(error_layername = ERROR_LAYER_NAME):
    doc = helpercad.get_cad_active_doc()
    if doc is None:
        return
    try:
        helpercad.add_layer(doc,error_layername)
        selected_bar_info = helpercad.get_barinfo_list(doc)
        barmarks = set([x.attributes['BAR_MARK'] for x in selected_bar_info])
        grouped_bars = {}
        for key in barmarks:
            grouped_bars[key] = [x for x in selected_bar_info if x.attributes['BAR_MARK'] == key]
            check_equality = BarInfoBlock.check_barmark_equality(grouped_bars[key])
            for bar in grouped_bars[key]:
                bar.has_problem = not check_equality
        centers = [x.insertion_point for x in selected_bar_info if x.has_problem]
        helpercad.draw_circles(doc,ERROR_LAYER_NAME,centers,ERROR_CIRCLE_RADIUS)
    finally:
        doc = None


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