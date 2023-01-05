import pptxProcessor


input_pptx = r"C:\Users\ssellars\Documents\PPTX\testingPlayground\examples\test.pptx"
output_pptx = r"C:\Users\ssellars\Documents\PPTX\testingPlayground\examples\output.pptx"
context = {
    "first_name" : "John",
    "relationship_related_one" : [
        {
            "first_field": "one",
            "second_field" : "one",
            "third_field" : "one"
        },
        {
            "first_field": "two",
            "second_field" : "two",
            "third_field" : "two"
        },
        {
            "first_field": "three",
            "second_field" : "four",
            "third_field" : "five"
        }
    ]
    }


pptxProcessor.templatePptx(input_pptx, context, output_pptx, "$")

pptxProcessor.batchTool(r"C:\Users\ssellars\Documents\PPTX\testingPlayground\examples", output_pptx).combine_slides()






