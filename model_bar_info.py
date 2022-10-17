




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
