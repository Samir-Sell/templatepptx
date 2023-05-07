import templatepptx
import os


input_pptx = os.path.join(os.path.dirname(os.path.realpath(__file__)), r"examplePresentations\ExampleTables.pptx")
output_pptx = os.path.join(os.path.dirname(os.path.realpath(__file__)), r"examplePresentations\ExampleTablesOutput.pptx")
context = {
    "random_key" : "random_value",
         "relationship_test" : [ 
             { "id" : "1", "first_name" : "Duncan", "last_name" : "Junior" }, 
             { "id" : "2", "first_name" : "Jessica", "last_name" : "Jones" } 
          ],
          # Same relationship as a above but with the name "people" instead of "test"
         "relationship_people" : [ 
             { "id" : "3", "first_name" : "Duncan", "last_name" : "Junior" }, 
             { "id" : "4", "first_name" : "Jessica", "last_name" : "Jones" } 
          ]
    }


templatepptx.templatePptx(input_pptx, context, output_pptx, "$").parse_template_pptx()







