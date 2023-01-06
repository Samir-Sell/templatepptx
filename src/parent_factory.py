class parentProcessor():
    
    def __init__(self, shape, context, slide_number, special_character):
        
        self._shape = shape
        self._context = context
        self._slide_number = slide_number
        self._special_character = special_character

    @property
    def context(self):
        return self._context

    @context.setter
    def context(self, context):
        self._context = context