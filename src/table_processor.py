from src.parent_processor import ParentProcessor
from copy import deepcopy
from pptx.table import _Cell
import warnings
from src.text_processor import TextProcessor

class TableProcessor(ParentProcessor):

    def __init__(self, shape, context, slide_number, special_character):
        super().__init__(shape, context, slide_number, special_character)

    
    def _remove_row(self, table, row_num):

        """
        Description: Remove a row from a PPTX table that is found in a pptx shape object

        @input table: A table from a PPTX
        @input row_num: An integer containing the row number to be deleted. negative numbers can be used to index 
        backwords
        """
        table._tbl.remove(table._tbl.tr_lst[row_num])

    def _add_row(self, table, relationship_class, record_count):

        """
        Description: Add a row to a PPTX table using a relationship specified in the contect dictionary through dot notation ex. (relaltionship_name.field_value)

        @input table: A table from a PPTX
        @input context: A dictionary containing all of the data that is fed into the template. It contains the data 
        and the magic keywords.
        @input relationship_class: The name of the relationship class
        @input record_count: An integer specifying which item in the related records context relationship value 
        should be inserted into the cells. AKA: This selects a single dictionary from the list of dictionaries.

        """
        
        # duplicating last row of the table as a new row to be added
        new_row = deepcopy(table._tbl.tr_lst[1]) 
        # Loop through cells in new row
        for tc in new_row.tc_lst:
            # Get cell
            cell = _Cell(tc, new_row.tc_lst)
            # Get cell text
            cell_text = cell.text
            # Process and get rel name and field in rel
            cell_value = cell_text.replace(self._special_character, "")
            cell_keys = cell_value.split(".")
            
            cell_field = cell_keys[1] # Get relationship name
            cell_field = cell_field.strip('\n') # Strip out newlines if needed

            # Replace the placeholder formatting text with text from context
            cell_text = cell_text.replace(cell_text, self._context[relationship_class][record_count][cell_field])
            cell.text = cell_text        
        table._tbl.append(new_row) #Append to existing table
        
    def process_table(self):

        '''
        Description: Process a table and replace every value or populate a table based on a relationship
        
        @input shape: A container from PPTX Python called shape which contains paragraphs and text
        @input context: A dictionary containing all of the data that is fed into the template. It contains the data 
        and the magic keywords.
        @input slide_number: The slide number index
        '''

        try:
            table_cells = self._shape.table.iter_cells()
            for cell in table_cells:
                if (cell.text.find("relationship")) != -1:
                    cleaned_cell = (cell.text).replace(self._special_character, "")
                    relationship_class = (cleaned_cell).split(".")[0]
                    rel_class_key = self._context.get(relationship_class)
                    if rel_class_key: 
                        for row in range(len(self._context[relationship_class])):
                            self._add_row(self._shape.table, relationship_class, row)
                        self._remove_row(self._shape.table, 1)
                        break
                    else:
                        warnings.warn(f"Relationship link for {relationship_class} does not exist.")
                        raise KeyError(relationship_class)
                for p in cell.text_frame.paragraphs:
                    TextProcessor(self._shape, self._context, self._slide_number, self._special_character)._replace_runs(p)
        except Exception as e:
            warnings.warn(f"Table failed to be populated due to {e}")