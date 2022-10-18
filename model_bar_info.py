




class BarInfoBlock:
    def __init__(self) -> None:
        self.name = ''
        self.handle = ''
        self.id = ''
        self.insertion_point = []
        self.attributes = {}
        self.has_problem = False

    def __eq__(self, __o: object) -> bool:
        if self.attributes['BAR_MARK'] != __o.attributes['BAR_MARK']:
            return False
        if self.attributes['BAR_SHAPE'] != __o.attributes['BAR_SHAPE']:
            return False
        if self.get_bar_dia() != __o.get_bar_dia():
            return False
        if self.get_total_length() != __o.get_total_length():
            return False
        return True


    def get_image_filename (self):
        long_name = self.attributes['BAR_SHAPE']
        return 'BAR-' + str(long_name.split('-')[1]) + '.WMF'


    def get_bar_dia(self):
        info = self.attributes['BAR_INFO']
        diameter = info.split('@')[0].split('T')[1]
        return diameter
    def get_total_length(self):
        info = self.attributes['BAR_INFO']
        length = info.split('x')[1]
        return length

    def get_bar_count(self):
        info = self.attributes['BAR_INFO']
        count = info.split('-')[0]
        result = 0
        if '+' in count:
            x = int(count.split('+')[0])
            y = int(count.split('+')[1])
            result = x + y
        elif '*' in count:
            x = int(count.split('*')[0])
            y = int(count.split('*')[1])
            result = x * y
        else:
            result = int(count)
        return result

    def check_barmark_equality(bar_list):
        check_bar = bar_list[0]
        for bar in bar_list[1:]:
            if bar != check_bar:
                return False
        return True

    def get_total_count(bar_list):
        bar_count = [x.get_bar_count() for x in bar_list]
        bar_count = sum(bar_count)
        return bar_count

    def group_barlist_by_barmark(bar_list):
        barmarks = set([x.attributes['BAR_MARK'] for x in bar_list])
        grouped_bars = {}
        check_equality = False
        for key in barmarks:
            grouped_bars[key] = [x for x in bar_list if x.attributes['BAR_MARK'] == key]
            check_equality = BarInfoBlock.check_barmark_equality(grouped_bars[key])
        if check_equality:
            return grouped_bars
        else:
            print('<<*****>>\tSelected Bar Info blocks are not equals\t<<*****>>')