[![Downloads](https://pepy.tech/badge/templatepptx)](https://pepy.tech/project/templatepptx)

# Description

Use PowerPoint templates to generate PowerPoint files based on PowerPoint templates. The PowerPoints are generated on the fly using "magic words". Magic words are specified by using the `$` sign symbol. You can specify magic words in PowerPoint templates by wrapping the word like `$this$`. Pictures can also be used as templates and are specified by defining the key words in the alt text of the picture. This tool is simple to run and setup. 

## How to Install 
`pip install templatepptx`

The data is populated by using a "context" object. A context object is a dictionary which contains the keywords and thier values that are used to populate the powerpoint. Additionally, tables can be populated with an unlmited number of related data by specifying a list of dictionaries in your context. A related table variable is specified in the template by adding the prefix "relationship_" to the front of the key. Please observe the following example of a context object below.

## Parsing Quick Start

To run this tool you will need a template PowerPoint that contains slides that have magic keywords. You will need a context file with the key words and you will need a valid PPTX file path for the output.

```
import templatepptx

input_pptx = "path//to//input.pptx"
output_pptx = "path//to//output.pptx"
context = {
    "first_name" : "John",
    "last_name" : "Smith",
    "language" : "Python",
    "title" : "PPT Tool",
    "italics" : "italics",
    "feeling" : "happy",
    "example_photo" : "path/to/example/photo.png"
    "relationship_name" : [ # This key contains the list which can contain an unlimited amount of records to populate a table.
    {
        "id" : "1",
        "first_name" : "Duncan",
        "last_name" : "Junior"},
    {
        "id" : "2",
        "first_name" : "Jessica",
        "last_name" : "Jones"}]
    }

# Read in PowerPoint and Context. Also assign what the special character is.
powerpoint = templatepptx.templatePptx(input_pptx, context, output_pptx, "$")

# Parses and exports the PowerPoint with filled out values and pictures
powerpoint.parse_template_pptx()

```

## Combining Slides Quickstart

You can generate many similar output products from a group of templates and then combining these outputs into one final product. There is an automated function built into this module which permits you to point to a whole direct, scrape all of the .pptx files and then combine them into one .pptx file. 

```
import templatepptx
in_dir = "path//to//input_dir"
out_combined = "path//to//combined_output.pptx"
templatepptx.batchTool(in_dir, out_combined).combine_slides():
```

# Documentation

## templatepptx module

##### Class `templatepptx.templatePptx(ppt, context, output_path, special_character="$")`

*Description:*
Initializes templatePptx currently provides the ability to completely parse through a template PowerPoint and replace the magic words, tables and pictures with the desired data from the context.

*Class Parameters:*
-   `ppt` : File path to template PowerPoint to parse (This file must exist). Required.
-   `context` : Dictionary containing key pair values for magic words and their new desired value. Required.
-   `output_path` : File path to the location where parsed PowerPoint will be written to. Required.
-   `special_character` : Special character which is wrapped around key words. The special character is not required and defaults to `$`. Example: `$this$`. If dollar signs do not suffice, it can be changed. Optional.

*Methods:*
-   `templatepptx.templatePptx.parse_template_pptx()` Runs method from templatePptx to parse the template.

*Properties:*
-   `context` Getter and Setter to change and view Context on the fly


*Example:*
```
import templatepptx

# Initialize templatePptx class
ppt = "path/to/template.pptx"
context = {"template_word" : "desired_new_value",
            "alt_text_key" : "path/to/image.jpg"}
output_path = "path/to/new/output.pptx"
powerpoint_template = templatepptx.templatePptx(ppt, context, output_path, special_character="$")

# Parse template
powerpoint_template.parse_template_pptx()
```

##### Class `templatepptx.batchTool(pptx_dir, output_pptx)`

*Description:*
Initalizes the batch tool to combine PowerPoints. 

*Class Parameters:*
-   `pptx_dir` : Directory path to the directory containing multiple PowerPoint files to be combined.
-   `output_pptx` : File path to the desired output location of the combined PowerPoint.

*Methods:*
-   `templatepptx.batchTool.combine_slides()` Runs the method to combine slides and output all slides into one PPTX. 
    - `is_numeric` : Boolean which defaults to True. Combine slides will attempt to combine slides in the correct numerical order that contain only numeric digits such as 1, 2 or 3. For examples, the following directory containing 1.pptx, 4.pptx and 2.pptx will be combined using slides from 1 first, 2 second and 4 last.
    - `specify_master` : A file path which specifies if a blank master deck exists. Defaults to None and creates a blank template for you. Allows for slide masters to be used which contain certain themes that will persist when combining slides. Text and images on a slide master will NOT be parsed and will remain intact. ONLY blank slide templates are used to create and copy PowerPoint templates, therefore only the blank Slide Master slide will be seen in the end product. 

## Future Planned Features
- ArcGIS Feature Service Support (Ask as needed)
- MSSQL support


## Example

Example input slides.
![input slide 1 example](img/in1.PNG)
![input slide 2 example](img/in2.PNG)

Example output slides.
![output slide 1 example](img/out1.PNG)
![output slide 2 example](img/out2.PNG)
