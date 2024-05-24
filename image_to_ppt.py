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
                changed_files.add(file_path.strip())
    
    # Include changes from the latest commit
    if len(commit_ids) > 0:
        command = f"git diff --name-status {commit_ids[-1]}^!"
        diff_output = subprocess.check_output(command, shell=True).decode().splitlines()
        
        for line in diff_output:
            status, file_path = line.split(maxsplit=1)
            if file_path.startswith(image_dir):
                changed_files.add(file_path.strip())
    
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
        image_width = (slide_width - Inches(1)) / 2  # Leave some space between images
        image_height = slide_height - Inches(1)  # Leave some space at top and bottom

        left1 = Inches(0.25)
        top1 = Inches(0.5)
        left2 = left1 + image_width + Inches(0.5)
        
        # Add the first image and label
        if i < len(image_paths):
            image_path1 = image_paths[i]
            print(f"Adding image: {image_path1}")  # Debug print
            slide.shapes.add_picture(image_path1, left1, top1, image_width, image_height)
            
            textbox = slide.shapes.add_textbox(left1, Inches(0.0), image_width, Inches(0.5))
            text_frame = textbox.text_frame
            p = text_frame.add_paragraph()
            p.text = "Deleted"
            p.font.size = Pt(18)

        # Add the second image and label if it exists
        if i + 1 < len(image_paths):
            image_path2 = image_paths[i + 1]
            print(f"Adding image: {image_path2}")  # Debug print
            slide.shapes.add_picture(image_path2, left2, top1, image_width, image_height)
            
            textbox = slide.shapes.add_textbox(left2, Inches(0.0), image_width, Inches(0.5))
            text_frame = textbox.text_frame
            p = text_frame.add_paragraph()
            p.text = "Updated"
            p.font.size = Pt(18)
    
    presentation.save(ppt_path)

# Get the list of changed image files
changed_images = get_changed_files(n=3)
print(f"Changed images: {changed_images}")  # Debug print

# Create a PPT with the changed images
create_ppt_with_images(changed_images)
