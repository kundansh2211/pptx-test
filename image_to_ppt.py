import os
from pptx import Presentation
from pptx.util import Inches

# Define the directory where images are stored
image_dir = 'images'

# List of images to be converted to ppt
image_filenames = ['image.jpg', 'image2.jpg', 'image3.jpg', 'image4.jpg']

presentation = Presentation()

margin_inch = 0.5 

# Convert margin to pptx units (EMU: English Metric Unit, 1 inch = 914400 EMUs)
margin = Inches(margin_inch)

# Loop through the image filenames and add slides with images
for i in range(0, len(image_filenames), 2):
    # Create a slide with blank layout
    slide_layout = presentation.slide_layouts[6]
    slide = presentation.slides.add_slide(slide_layout)

    # Define the height and width of the slide
    slide_width = presentation.slide_width
    slide_height = presentation.slide_height

    # Calculate the width and height for the images
    image_width = (slide_width - 3 * margin) / 2  # Two images side by side with margins
    image_height = slide_height - 2 * margin      # Image height considering top and bottom margins

    # Positions for the first image
    left1 = margin
    top1 = margin

    # Add the first image to the slide if available
    if i < len(image_filenames):
        image_path1 = os.path.join(image_dir, image_filenames[i])
        if os.path.isfile(image_path1):
            slide.shapes.add_picture(image_path1, left1, top1, image_width, image_height)

    # Positions for the second image
    left2 = 2 * margin + image_width
    top2 = margin

    # Add the second image to the slide if available
    if i + 1 < len(image_filenames):
        image_path2 = os.path.join(image_dir, image_filenames[i + 1])
        if os.path.isfile(image_path2):
            slide.shapes.add_picture(image_path2, left2, top2, image_width, image_height)

# Save the presentation
presentation.save('PPT/image_ppt.pptx')
