import os
import subprocess
from pptx import Presentation
from pptx.util import Inches, Pt

# Define the directory where images are stored
image_dir = 'images'

def get_changed_files(n=3):
    # Get the last n commit IDs
    command = f"git log -n {n} --pretty=format:%H"
    commit_ids = subprocess.check_output(command, shell=True).decode().split()
    changed_files = set()
    
    for i in range(len(commit_ids) - 1):
        commit_id = commit_ids[i]
        prev_commit_id = commit_ids[i + 1]
        
        # Get the diff of files between each pair of commits
        command = f"git diff --name-status {commit_id} {prev_commit_id}"
        diff_output = subprocess.check_output(command, shell=True).decode().splitlines()
        
        for line in diff_output:
            status, file_path = line.split(maxsplit=1)
            if file_path.startswith(image_dir):
                changed_files.add(file_path)
    
    # Include changes from the latest commit
    if len(commit_ids) > 0:
        command = f"git diff --name-status {commit_ids[-1]}^!"
        diff_output = subprocess.check_output(command, shell=True).decode().splitlines()
        
        for line in diff_output:
            status, file_path = line.split(maxsplit=1)
            if file_path.startswith(image_dir):
                changed_files.add(file_path)
    
    return list(changed_files)

def create_ppt_with_images(image_paths, ppt_path='PPT/image_ppt.pptx'):
    presentation = Presentation()
    
    for i in range(0, len(image_paths), 2):
        # Create a slide with blank layout
        slide_layout = presentation.slide_layouts[6]
        slide = presentation.slides.add_slide(slide_layout)

        # Define the dimensions for the images
        slide_width = presentation.slide_width
        slide_height = presentation.slide_height
        image_width = slide_width / 2 - Inches(0.5)  # Leave some space between images
        image_height = slide_height - Inches(1)  # Leave some space at top and bottom

        left1 = Inches(0.25)
        top1 = Inches(0.5)
        left2 = slide_width / 2 + Inches(0.25)
        
        # Add the first image and label
        if i < len(image_paths):
            image_path1 = image_paths[i]
            slide.shapes.add_picture(image_path1, left1, top1, image_width, image_height)
            
            textbox = slide.shapes.add_textbox(left1, Inches(0.0), image_width, Inches(0.5))
            text_frame = textbox.text_frame
            p = text_frame.add_paragraph()
            p.text = "Deleted"
            p.font.size = Pt(18)

        # Add the second image and label if it exists
        if i + 1 < len(image_paths):
            image_path2 = image_paths[i + 1]
            slide.shapes.add_picture(image_path2, left2, top1, image_width, image_height)
            
            textbox = slide.shapes.add_textbox(left2, Inches(0.0), image_width, Inches(0.5))
            text_frame = textbox.text_frame
            p = text_frame.add_paragraph()
            p.text = "Updated"
            p.font.size = Pt(18)
    
    presentation.save(ppt_path)

# Get the list of changed image files
changed_images = get_changed_files(n=3)

# Create a PPT with the changed images
create_ppt_with_images(changed_images)








# import os
# from pptx import Presentation
# from pptx.util import Inches

# # Define the directory where images are stored
# image_dir = 'images'

# # List of images to be converted to ppt
# image_filenames = ['image.jpg', 'image2.jpg', 'image3.jpg', 'image4.jpg']

# presentation = Presentation()

# margin_inch = 0.5 

# # Convert margin to pptx units (EMU: English Metric Unit, 1 inch = 914400 EMUs)
# margin = Inches(margin_inch)

# # Loop through the image filenames and add slides with images
# for i in range(0, len(image_filenames), 2):
#     # Create a slide with blank layout
#     slide_layout = presentation.slide_layouts[6]
#     slide = presentation.slides.add_slide(slide_layout)

#     # Define the height and width of the slide
#     slide_width = presentation.slide_width
#     slide_height = presentation.slide_height

#     # Calculate the width and height for the images
#     image_width = (slide_width - 3 * margin) / 2  # Two images side by side with margins
#     image_height = slide_height - 2 * margin      # Image height considering top and bottom margins

#     # Positions for the first image
#     left1 = margin
#     top1 = margin

#     # Add the first image to the slide if available
#     if i < len(image_filenames):
#         image_path1 = os.path.join(image_dir, image_filenames[i])
#         if os.path.isfile(image_path1):
#             slide.shapes.add_picture(image_path1, left1, top1, image_width, image_height)

#     # Positions for the second image
#     left2 = 2 * margin + image_width
#     top2 = margin

#     # Add the second image to the slide if available
#     if i + 1 < len(image_filenames):
#         image_path2 = os.path.join(image_dir, image_filenames[i + 1])
#         if os.path.isfile(image_path2):
#             slide.shapes.add_picture(image_path2, left2, top2, image_width, image_height)

# # Save the presentation
# presentation.save('PPT/image_ppt.pptx')
