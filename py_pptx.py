import collections 
import collections.abc
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

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


Image1 = "/home/hunter/Documents/Workspace/Python_code/py_pptx/image1.jpg"
Image2 = "/home/hunter/Documents/Workspace/Python_code/py_pptx/image2.jpg"

generate_ppt(Image1, Image2, 'My Presentation Title',"presentation")
# Open the PowerPoint file
# prs = Presentation("presentation.pptx")
# open("presentation.pptx", "r")