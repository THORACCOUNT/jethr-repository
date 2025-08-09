from presentation import JetHRPresentation
import pandas

OUTPUT_PATH = r"../data/output/slides_presentation.pptx"
INPUT_PATH = "../data/input/slides_content.csv"
# Create object JetHRPresentation to produce pptx presentation and save it onto the output path
presentation = JetHRPresentation(OUTPUT_PATH)
# Read the csv containing data to fill in the slide using pandas
slide_content = pandas.read_csv(INPUT_PATH)
# Extract Series of title and content
title_list = slide_content.title
content_list = slide_content.content
# By using pandas shift method, reduce title Series to only different items
title_list_diff = title_list[title_list != title_list.shift()]
# We call add_title_slide method to create first slide - to do this we need to remove whitespaces in (title, content) values
presentation.add_title_slide(title_list[0].strip(), content_list[0].strip())
# We convert Series of different titles to a list and remove first element, already dealt with, for commodity
title_list_diff_1 = title_list_diff.to_list()
title_list_diff_1.remove(title_list[0])
# For each title in our selection of unique ones we extract a new list of contents from our DataFrame only related to it
for title in title_list_diff_1:
    content_related_to_title = slide_content[title_list == title].content
    # We will use join to concatenate strings contained in content_related_to_title
    content_string = "\n".join(content for content in content_related_to_title)
    # We add title, and it's related to content to the slide
    presentation.add_content_slide(title, content_string)
# Finally we call our method to save the presentation
presentation.save_presentation()
