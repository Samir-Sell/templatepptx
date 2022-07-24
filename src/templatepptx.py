"""
 ---------------------------------------------------------------------------
 Name        : main.py
 Description : Tool that uses template PowerPoint files to generate new PowerPoint files bases on dictionary values and "magic words" 
 Created     : 18-July-2022
 Author      : Samir Sellars
 Usage       : 

 Notes:
"""

# Import used to fix if on Python 3.10 according to: https://github.com/scanny/python-pptx/issues/762
import collections.abc
from concurrent.futures import process
from multiprocessing.sharedctypes import Value

from pptx import Presentation
from pptx.table import _Cell
from copy import deepcopy
import os
import json
import sys
import warnings
import glob
import win32com.client

#Global Constants
SPECIAL_CHARACTER = "$"
MAIN_PATH = os.path.realpath(__file__)
SCRIPT_DIR = os.path.dirname(MAIN_PATH)


def error(e, kill=False, message=None):
    
    '''
    Description: Prints errors onto screen

    @input e: Exception Object
    @input kill: Bool stating if program should crash
    @input message: String containing helpful info the user
    '''

    print("ERROR OCCURRED:")
    if message is not None:
        print(message)
    print(f"Error: {e}")
    exc_type, exc_obj, exc_tb = sys.exc_info()
    fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
    print(f"Error Type: {exc_type}, File Name: {fname}, Line Number: {exc_tb.tb_lineno}")
    if kill == True:
        print("Exiting Program...")
        sys.exit()

def out_warning(w):

    '''
    Description: Prints warning onto screen

    @input w: String containing warning message
    '''

    warnings.warn(w)


def formatting_check(run, last_run) -> bool:

    """
    description: Checks for the formatting of the different text runs and compares them

    @input run: Run object that contains formatting and text for a small segment of text
    @input last_run: Run object that contains formatting and text for a small segment of text from the last run
    
    @output: Boolean containing True if current run matches last run
    """
    
    # Compare two runs formatting
    if last_run == None:
        return False
    run_format = {'size':run.font.size,
                'bold':run.font.bold,
                'underline':run.font.underline,
                'italic':run.font.italic,
                'name' :run.font.name}
    last_run_format = {'size':last_run.font.size,
                'bold':last_run.font.bold,
                'underline':last_run.font.underline,
                'italic':last_run.font.italic,
                'name' :last_run.font.name}
    if run_format == last_run_format:
        return True
    else:
        return False
   
def replace_runs(p, slide_number, context):
    
    """
    Description: Function to replace runs for each paragraph object.
    
    @input p: Paragraph object from pptx that contains multiple runs
    @input slide_number: integer containing the slide number
    @input context: A dictionary containing all of the data that is fed into the template. It contains the data 
    and the magic keywords.
    """

    try: 
        last_run = None
        # Loop through every run in a paragraph obj
        for run in p.runs:
            # Check if the current run and last have the same formatting. If they do, combine the runs
            if formatting_check(run, last_run):
                combined_text = last_run.text + run.text
                for key in context:
                    to_find = add_special(key, SPECIAL_CHARACTER)
                    combined_text = combined_text.replace(str(to_find), str(context[key]))
                run.text = combined_text
                # Remove the current run after it has been combined with the last run
                run_to_remove = last_run._r
                run_to_remove.getparent().remove(run_to_remove)
            # If formatting does not align
            else:
                # Replace values in the rund
                text = run.text
                for key in context:
                    to_find = add_special(str(key), SPECIAL_CHARACTER)
                    text = text.replace(str(to_find), str(context[key]))
                run.text = text
            last_run = run
    except Exception as e:
        error(e, kill=True, message=f"A problem occurred in a run on slide number {slide_number}")


def add_special(key, SPECIAL_CHARACTER) -> str:
    
    """
    Description: Adds a 'magic' character to either side of the dictionary key being handed to the function

    @input key: A key from the context dictionary that acts as a magic word in the PPTX template
    @input SPECIAL_CHARACTER: A character that is added to the front and the end of the key
    """
    return f"{SPECIAL_CHARACTER}{key}{SPECIAL_CHARACTER}"


def remove_row(table, row_num):
    
    """
    Description: Remove a row from a PPTX table that is found in a pptx shape object

    @input table: A table from a PPTX
    @input row_num: An integer containing the row number to be deleted. negative numbers can be used to index 
    backwords
    """
    table._tbl.remove(table._tbl.tr_lst[row_num])

def add_row(table, context, relationship_class, record_count):

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
        cell_value = cell_text.replace(SPECIAL_CHARACTER, "")
        cell_keys = cell_value.split(".")
        
        cell_field = cell_keys[1] # Get relationship name
        cell_field = cell_field.strip('\n') # Strip out newlines if needed

        # Replace the placeholder formatting text with text from context
        cell_text = cell_text.replace(cell_text, context[relationship_class][record_count][cell_field])
        cell.text = cell_text        
    table._tbl.append(new_row) #Append to existing table
    
def process_table(shape, context, slide_number):

    '''
    Description: Process a table and replace every value or populate a table based on a relationship
    
    @input shape: A container from PPTX Python called shape which contains paragraphs and text
    @input context: A dictionary containing all of the data that is fed into the template. It contains the data 
    and the magic keywords.
    @input slide_number: The slide number index
    '''

    try:
        table_cells = shape.table.iter_cells()
        for cell in table_cells:
            if (cell.text.find("relationship")) != -1:
                cleaned_cell = (cell.text).replace(SPECIAL_CHARACTER, "")
                relationship_class = (cleaned_cell).split(".")[0]
                rel_class_key = context.get(relationship_class)
                if rel_class_key: 
                    for row in range(len(context[relationship_class])):
                        add_row(shape.table, context, relationship_class, row)
                    remove_row(shape.table, 1)
                    break
                else:
                    out_warning(f"Relationship link for {relationship_class} does not exist.")
                    raise KeyError(relationship_class)
            for p in cell.text_frame.paragraphs:
                replace_runs(p, slide_number, context)
    except Exception as e:
        error(e, kill=False, message=f"Table failed to be populated and is being skipped. This is not a fatal error but it will not populate your table on slide {slide_number}")


def replace_text(shape, context, slide_number):

    '''
    Description: Process text and replace every value
    
    @input shape: A container from PPTX Python called shape which contains paragraphs and text
    @input context: A dictionary containing all of the data that is fed into the template. It contains the data 
    and the magic keywords.
    @input slide_number: The slide number index
    '''

    if shape.has_text_frame:
        for key in context:
            if(shape.text.find(str(key)))!=-1:
                text_frame = shape.text_frame
                for p in text_frame.paragraphs:
                    replace_runs(p, slide_number, context)

def parse_template_pptx(ppt, context, output_path) -> Presentation:
    
    """
    Description: The parent function that parses the powerpoint into a PPTX Presentation and replaces magic words

    @input ppt: A file path to the template PPTX
    @input context: A dictionary containing all of the data that is fed into the template. It contains the data 
    and the magic keywords.

    @output ppt: A Python pptx Presentation object which contains all of the new changes
    """

    # Find template pptx
    try:
        ppt = Presentation(ppt)
    except Exception as e:
        error(e, kill=True, message="Could not find PowerPoint file")

    

    # Warn user if context obj is empty
    if context == {}:
        out_warning("Context file is empty")
    
    # Check if the context is a valid dictionary
    if not isinstance(context, dict):
        try:
            raise ValueError
        except Exception as e:
            error(e, kill=True, message="Your context is not a valid dictionary. Please check the parameter type.")
    
    # Check if the output is valid
    try:
        with open(output_path, 'w') as out_pptx:
            pass
    except Exception as e:
        error(e, kill=True, message="Cannot open a PPTX file at the desired output dir")



    # Loop through every shape element in each slide and replace template words with values from context
    for slide in ppt.slides:
        slide_number = (ppt.slides.index(slide)) + 1
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                replace_text(shape, context, slide_number)
                                    
            # If shape object has a table associated, process table
            #NOTE: relationship is a key word and is used to specify table relates                      
            if shape.has_table:
                process_table(shape, context, slide_number)                                
    
    try:
        ppt.save(output_path)
        print(f"PowerPoint generated at {output_path}")
        return output_path

    except Exception as e:
        error(e, kill=True, message="Could not write the end product powerpoint. Possibly due to PPTX being open or file path not existing.")    


def combine_slides(in_dir, out_dir):

    '''
    Description: Combine multiple PPTX files into one
    
    @input in_dir: A directory path string containing 2 or more .pptx files
    @input out_dir: A file path string pptx file that will be the final output
    '''

    # Find all slides in the temp output dir
    pres = glob.glob(os.path.join(in_dir,"*.pptx"))
    if pres == []:
        out_warning("No PowerPoints were found in the specified directory")

    # Launches PowerPoint and opens first PPT
    try:
        ppt_instance = win32com.client.Dispatch('PowerPoint.Application')
        prs = ppt_instance.Presentations.open(os.path.abspath(pres[0]), True, False, False)

        # For the other PPTX files, insert slides from other slides
        for i in range(1, len(pres)):
            prs.Slides.InsertFromFile(os.path.abspath(pres[i]), prs.Slides.Count)
        prs.SaveAs(os.path.abspath(out_dir))
        prs.Close()
    except Exception as e:
        error(e, kill=False, message="Failed to combine powerpoints. It is likely no PPTX were found. This functionality also requires PowerPoint to be installed on a Windows Machine.")
        pass
    