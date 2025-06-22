from pptx.shapes.autoshape import Shape

class ParentProcessor():
    
    def __init__(self, shape: Shape, context: dict, slide_number: int, special_character: str):
        
        self._shape = shape
        self._context = context
        self._slide_number = slide_number
        self._special_character = special_character

    @property
    def context(self) -> dict:
        return self._context

    @context.setter
    def context(self, context) -> None:
        self._context = context