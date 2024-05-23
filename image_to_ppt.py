import os
from pptx import Presentation
from pptx.util import Inches

# Define the directory where images are stored
image_dir = 'images'

# List of images to be converted to ppt (ensure an even number of images)
image_filenames = ['image1.jpg', 'image2.jpg', 'image3.jpg', 'image4.jpg']

# Create a presentation object
presentation = Presentation()

# Loop through the image filenames in pairs and add slides with two images each
for i in range(0, len(image_filenames), 2):
    # Construct the full paths to the image files
    image_path1 = os.path.join(image_dir, image_filenames[i])
    image_path2 = os.path.join(image_dir, image_filenames[i + 1])

    # Check if the files exist
    if not os.path.isfile(image_path1):
        print(f"File not found: {image_path1}")
        continue
    if not os.path.isfile(image_path2):
        print(f"File not found: {image_path2}")
        continue

    try:
        # Create a slide with blank layout
        slide_layout = presentation.slide_layouts[6]
        slide = presentation.slides.add_slide(slide_layout)

        # Define the height of the images to maintain aspect ratio
        slide_width = presentation.slide_width
        slide_height = presentation.slide_height

        # Define the width of each image (half the slide width)
        image_width = slide_width / 2
        image_height = slide_height

        # Add the first image to the left half of the slide
        left1 = 0
        top1 = (slide_height - image_height) / 2  # Center vertically
        slide.shapes.add_picture(image_path1, left1, top1, width=image_width)

        # Add the second image to the right half of the slide
        left2 = image_width
        top2 = (slide_height - image_height) / 2  # Center vertically
        slide.shapes.add_picture(image_path2, left2, top2, width=image_width)
        
    except Exception as e:
        print(f"Error adding images {image_path1} and {image_path2}: {e}")

# Save the presentation
presentation.save('PPT/image_ppt.pptx')
