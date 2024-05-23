import os
from pptx import Presentation

# Define the directory where images are stored
image_dir = 'images'

# List of images to be converted to ppt
image_filenames = ['image.jpg', 'image2.jpg', 'image3.jpg', 'image4.jpg']

# Create a presentation object
presentation = Presentation()

# Loop through the image filenames and add slides with images
for filename in image_filenames:
    # Construct the full path to the image file
    image_path = os.path.join(image_dir, filename)

    # Check if the file exists
    if not os.path.isfile(image_path):
        print(f"File not found: {image_path}")
        continue

    try:
        # Create a slide with blank layout
        slide_layout = presentation.slide_layouts[6]
        slide = presentation.slides.add_slide(slide_layout)

        # Define the height and width of the image
        slide_width = presentation.slide_width
        slide_height = presentation.slide_height
        image_width = slide_width / 2  # Adjusted to fit two images per slide
        image_height = slide_height

        # Calculate the position to place the images side by side
        left1 = (slide_width / 2 - image_width) / 2
        left2 = slide_width / 2 + (slide_width / 2 - image_width) / 2
        top = (slide_height - image_height) / 2

        # Add the images to the slide
        slide.shapes.add_picture(image_path, left1, top, image_width, image_height)
        # Assuming you have a second image for each slide for demonstration
        # This is just an example, modify as needed
        if image_filenames.index(filename) + 1 < len(image_filenames):
            image_path2 = os.path.join(image_dir, image_filenames[image_filenames.index(filename) + 1])
            slide.shapes.add_picture(image_path2, left2, top, image_width, image_height)
    except Exception as e:
        print(f"Error adding image {image_path}: {e}")

# Ensure the output directory exists
output_dir = 'PPT'
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# Save the presentation
presentation.save(os.path.join(output_dir, 'image_ppt.pptx'))
