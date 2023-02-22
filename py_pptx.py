import collections 
import collections.abc
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import openpyxl
from pptx.enum.chart import XL_CHART_TYPE


Image1 = "/home/hunter/Documents/Workspace/Python_code/py_pptx/image1.jpg"
Image2 = "/home/hunter/Documents/Workspace/Python_code/py_pptx/image2.jpg"

def generate_ppt(image1_path, image2_path, title,pres_name):
    
    """Generate a PowerPoint presentation with two slides, each displaying one of the input images and the specified title at the bottom left in bold white text.

    Args:
        image1_path (str): The file path of the first image to be inserted into the presentation.
        image2_path (str): The file path of the second image to be inserted into the presentation.
        title (str): The title to be displayed at the bottom left of each slide.

    Returns:
        None.
    
    Raises:
        FileNotFoundError: If either of the image files does not exist.

    Example:
        generate_ppt('image1.jpg', 'image2.jpg', 'My Presentation Title')
    """
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

    prs.save(f"{pres_name}.pptx")

def append_slides(pptx_file):
    """Append 12 slides with a background image and a title at the header text box with some text in it and a small title.

    Args:
        pptx_file (str): The file path of the PowerPoint file to which to append the slides.

    Returns:
        None.

    Example:
        append_slides('example.pptx')
    """
    # Create presentation object
    prs = Presentation(pptx_file)

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
    # Save presentation
    prs.save(pptx_file)

def copy_charts_to_powerpoint(excel_file, pptx_file):
    """Copy a chart from each sheet of an Excel workbook (starting from sheet 5 until the last one) and paste it into a given PowerPoint file (starting from slide 3).

    Args:
        excel_file (str): The file path of the Excel workbook.
        pptx_file (str): The file path of the PowerPoint file to which to paste the charts.

    Returns:
        None.

    Example:
        copy_charts_to_powerpoint('data.xlsx', 'presentation.pptx')
    """
    import xlsxwriter
    from pptx import Presentation
    from pptx.util import Inches
    from tempfile import NamedTemporaryFile
    from PIL import Image

    # Open the Excel file
    workbook = xlsxwriter.Workbook('step_chart.xlsx')

    # Open the PowerPoint file
    prs = Presentation('presentation.pptx')

    # Loop through each sheet in the Excel file
    for sheet in workbook.worksheets():
        # Get the chart on the sheet
        chart = sheet.charts[0]

        # Get the image of the chart
        img_file = NamedTemporaryFile(delete=False)
        chart_img = chart.render()
        chart_img.save(img_file.name)

        # Add the image to the PowerPoint slide
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        left = top = Inches(1)
        slide.shapes.add_picture(img_file.name, left, top)

        # Clean up the temporary file
        img_file.close()

    # Save the PowerPoint file
    prs.save('presentation.pptx')




generate_ppt(Image1, Image2, 'My Presentation Title',"presentation")
append_slides("presentation.pptx")
copy_charts_to_powerpoint("step_chart.xlsx","presentation.pptx")

