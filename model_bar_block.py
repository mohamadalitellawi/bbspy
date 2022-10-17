class BarBlockData:
    def __init__(self) -> None:
        self.name = ''
        self.handle = ''
        self.id = ''
        self.dimensions = []
        self.insertion_point = []

    def get_original_name(self):
        long_name = self.name
        return long_name.split('XX')[0]

    def get_bar_shapename (self):
        long_name = self.get_original_name()
        return 'SHAPE-' + long_name.split('-')[1]

    def get_image_filename (self):
        long_name = self.get_original_name()
        return 'BAR-' + str(long_name.split('-')[1]) + '.WMF'