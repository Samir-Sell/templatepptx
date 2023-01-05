from parentFactory import parentProcessor
from copy import deepcopy
import warnings

class pictureProcessor(parentProcessor):

    def __init__(self, shape, context, slide_number, slide, special_character="$"):
        super().__init__(shape, context, slide_number, special_character)
        self._slide = slide

    
    def replace_picture(self):

        """
        Description: The function to replace an image in the PowerPoint Template
        @input shape: A container which has a image attribute associated
        @input context: A dictionary containing all of the data that is fed into the template. It contains the data 
        and the magic keywords.
        @input slide_number: integer containing the slide number
        @input slide: Slide object containing the shapes
        """
        # Get info about template picture in order to mimic it
        img_width = self._shape.width
        img_height = self._shape.height
        img_left = self._shape.left
        img_top = self._shape.top
        alt_text = self._get_alt_text()

        # Find matching picture if it exists
        alt_text_string = self._context.get(alt_text) # Outputs None if not valid
        
        # If found, remove the template picture and then add the new picture
        if alt_text_string != None:
            sp = self._shape._element # Get xml element
            sp.getparent().remove(sp) # Remove xml element
            self._slide.shapes.add_picture(image_file=alt_text_string, left=img_left, top=img_top, width=img_width, height=img_height)
            return alt_text_string
        else:
            warnings.warn(f"No image was found to be assoicated with this alt text. Template will remain in the PowerPoint. Alt Text: {alt_text} Slide Number: {self._slide_number}")
        



    def _get_alt_text(self):

        """
        Description: Gets alt text from an image

        @input shape: A container which has picture attribure associated with it
        @input slide_number: Used to troubleshoot upon failure

        @output context: A string containing the alt text
        """
        alt_text = None
        try:
            alt_text = self._shape.element.xpath("//p:pic/p:nvPicPr/p:cNvPr")[0].attrib["descr"]
        except Exception as e:
            warnings.warn(f"Error reading the alt text of picture on slide {self._slide_number}. Error: {e}")
        if alt_text == "" or alt_text == " " or alt_text == None:
            warnings.warning(f"Alt text is single space blank string or empty string and will not load any image. Slide: {self._slide_number}")
            return None

        return alt_text
