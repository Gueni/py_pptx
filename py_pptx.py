import collections 
import collections.abc
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import openpyxl
from pptx.enum.chart import XL_CHART_TYPE
import openpyxl
import os
# import win32com.client as win32
from pptx import Presentation
from pptx.util import Inches

Image1 = "/home/hunter/Documents/Workspace/Python_code/py_pptx/image1.jpg"
Image2 = "/home/hunter/Documents/Workspace/Python_code/py_pptx/image2.jpg"

def generate_ppt(image1_path, image2_path, title,pres_name):

    prs = Presentation()

    # set width and height to 16 and 9 inches.
    slide_width  = prs.slide_width = Inches(16)
    slide_height =prs.slide_height = Inches(9)

    # create first slide with image1
    slide1 = prs.slides.add_slide(prs.slide_layouts[6])
    pic1 = slide1.shapes.add_picture(image1_path, 0, 0, slide_width, slide_height)
    
    # create second slide with image2
    slide2 = prs.slides.add_slide(prs.slide_layouts[6])
    pic2 = slide2.shapes.add_picture(image2_path, 0, 0, slide_width, slide_height)

    # add title to both slides
    # for slide in [slide1, slide2]:
    title_shape = slide2.shapes.add_textbox(Inches(2), Inches(5.5), Inches(4), Inches(0.5))
    title_text_frame = title_shape.text_frame
    title_text_frame.text = title
    title_text_frame.paragraphs[0].font.bold = True
    title_text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
    title_text_frame.paragraphs[0].font.size = Pt(48)
    title_shape.top = prs.slide_height - title_shape.height - Inches(2.2)

    # Define background image path
    bg_image_path = 'Background.jpg'

    # Loop through and add 12 slides
    for i in range(1, 13):
        # Add slide
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        # Add background image
        slide.background.fill.solid()
        slide.background.fill.background()
        pic = slide.shapes.add_picture(bg_image_path, 0, 0, prs.slide_width, prs.slide_height)

        # Add header text box with title
        header_shape = slide.shapes.add_textbox(Inches(0), Inches(0.5), Inches(9), Inches(1))
        header_text_frame = header_shape.text_frame
        header_text_frame.text = f'Slide {i} Title'
        header_text_frame.paragraphs[0].font.size = Inches(0.5)
        header_text_frame.paragraphs[0].font.bold = True
        header_text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 0, 0)

        # Add body text box with small title
        body_shape = slide.shapes.add_textbox(Inches(0), Inches(1.5), Inches(9), Inches(1.5))
        body_text_frame = body_shape.text_frame
        body_text_frame.text = f'pptx_file (str): The file path of the \n PowerPoint file to which to append the slides.'
        body_text_frame.paragraphs[0].font.size = Inches(0.3)
        body_text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 0, 0)
        body_text_frame.paragraphs[0].font.bold = True
        # Adjust left position of text boxes to create indentation
        header_shape.left = Inches(0)
        body_shape.left = Inches(0)
    
    slide_width  = prs.slide_width = Inches(16)
    slide_height =prs.slide_height = Inches(9)
    # create second slide with image2
    slide2 = prs.slides.add_slide(prs.slide_layouts[6])
    pic2 = slide2.shapes.add_picture(Image1, 0, 0, slide_width, slide_height)
    prs.save(f"{pres_name}.pptx")


    # Open the Excel file and get its workbook object
    wb = openpyxl.load_workbook("step_chart.xlsx")
    # Define the path and name of the PowerPoint file
    pptx_path = f"{pres_name}.pptx"
    if os.name == 'nt':  # Check if the operating system is Windows
        # Open the PowerPoint file and get its application object
        ppt_app = win32.gencache.EnsureDispatch('PowerPoint.Application')
        ppt_app.Visible = True
        presentation = ppt_app.Presentations.Open(os.path.abspath(pptx_path))

    # Define the index of the slide to copy the charts to
    slide_index = 2  # Start from slide 2, since slide 1 might have a title or other content
    # Loop over the sheets of the Excel file
    for sheet_num in range(0,2):
        sheet = wb.worksheets[sheet_num]
        # Get the Drawing object for the sheet, if it exists
        drawing = sheet._drawing
        if drawing is not None:
            # Loop over the shapes in the Drawing object
            for shape in drawing:
                if shape.shape_type == 'chart':
                    # If the shape is a chart, check if it's a bar chart
                    if 'Bar' in shape.chart.chart_type:
                        if os.name == 'nt':  # Check if the operating system is Windows
                            # If it's a bar chart, copy it to the specified slide in the PowerPoint file (for Windows)
                            chart_copy = shape.copy_picture()
                            slide = presentation.Slides(slide_index)
                            slide.Shapes.Paste()
                        else:  # Otherwise, assume it's Linux
                            # If it's a bar chart, copy it to the specified slide in the PowerPoint file (for Linux)
                            chart = slide.shapes.add_chart(
                                chart_type=shape.chart.chart_type,
                                left=Inches(1),
                                top=Inches(1),
                                width=Inches(8),
                                height=Inches(4.5),
                            )
                            chart.chart.replace_data(shape.chart._chart_data)
        slide_index += 1  # Increment the slide index for the next sheet
        prs.save(pptx_path)
generate_ppt(Image1, Image2, 'My Presentation Title',"presentation")
