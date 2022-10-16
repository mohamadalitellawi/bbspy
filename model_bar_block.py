class BarBlockData:
    def __init__(self) -> None:
        self.name = ''
        self.handle = ''
        self.id = ''
        self.dimensions = []

    def get_bar_shape (self):
        long_name = self.name
        return 'SHAPE-' + long_name.split('-')[1]

    def get_image_filename (self):
        long_name = self.name
        return 'BAR-' + str(long_name.split('-')[1]) + '.WMF'