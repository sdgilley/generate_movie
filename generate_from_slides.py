import os
import subprocess
import re
from pptx import Presentation
from moviepy import ImageClip, concatenate_videoclips
from PIL import Image
import tempfile

def export_slides_as_images_libreoffice(pptx_file, output_dir="exported_slides"):
    """Try to export slides using LibreOffice (if available)"""
    os.makedirs(output_dir, exist_ok=True)
    
    try:
        # Try LibreOffice command
        cmd = [
            "soffice", "--headless", "--convert-to", "png", 
            "--outdir", output_dir, pptx_file
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
        
        if result.returncode == 0:
            print("Successfully exported slides using LibreOffice")
            return True
        else:
            print(f"LibreOffice export failed: {result.stderr}")
            return False
    except Exception as e:
        print(f"LibreOffice not available or failed: {e}")
        return False

def export_slides_as_images_powershell(pptx_file, output_dir="exported_slides"):
    """Try to export slides using PowerShell (Windows only)"""
    os.makedirs(output_dir, exist_ok=True)
    
    try:
        # PowerShell script to export PowerPoint slides
        ps_script = f'''
        $ppt = New-Object -ComObject PowerPoint.Application
        $ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse
        $presentation = $ppt.Presentations.Open("{os.path.abspath(pptx_file)}")
        
        for ($i = 1; $i -le $presentation.Slides.Count; $i++) {{
            $slide = $presentation.Slides.Item($i)
            $slide.Export("{os.path.abspath(output_dir)}\\slide_$i.png", "PNG")
        }}
        
        $presentation.Close()
        $ppt.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt)
        '''
        
        # Save PowerShell script to temp file
        with tempfile.NamedTemporaryFile(mode='w', suffix='.ps1', delete=False) as f:
            f.write(ps_script)
            ps_file = f.name
        
        # Execute PowerShell script
        cmd = ["powershell.exe", "-ExecutionPolicy", "Bypass", "-File", ps_file]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
        
        # Clean up temp file
        os.unlink(ps_file)
        
        if result.returncode == 0:
            print("Successfully exported slides using PowerShell")
            return True
        else:
            print(f"PowerShell export failed: {result.stderr}")
            return False
            
    except Exception as e:
        print(f"PowerShell export failed: {e}")
        return False

def export_slides_manual_instruction(pptx_file):
    """Provide manual instructions for exporting slides"""
    print(f"""
MANUAL EXPORT INSTRUCTIONS:
1. Open {pptx_file} in PowerPoint
2. Go to File > Export > Change File Type
3. Select PNG Portable Network Graphics Format
4. Click Save As
5. Choose a folder name like 'manual_slides'
6. PowerPoint will ask if you want to export "All Slides" - click "All Slides"
7. This will create individual PNG files for each slide
8. Come back and run this script again, it will look for the exported images
""")

def find_exported_slides():
    """Look for exported slide images in common locations"""
    possible_dirs = ["exported_slides", "manual_slides", "slide_exports", "slides"]
    
    for dir_name in possible_dirs:
        if os.path.exists(dir_name):
            png_files = [f for f in os.listdir(dir_name) if f.lower().endswith('.png')]
            if png_files:
                # Sort numerically instead of alphabetically
                import re
                def natural_sort_key(filename):
                    # Extract numbers from filename for proper sorting
                    numbers = re.findall(r'\d+', filename)
                    if numbers:
                        return int(numbers[0])  # Sort by first number found
                    return 0
                
                png_files.sort(key=natural_sort_key)
                return dir_name, png_files
    
    return None, []

# Load the PowerPoint presentation to get slide count and titles
pptx_file = "content_maintenance_process.pptx"
print(f"Loading PowerPoint presentation: {pptx_file}")
presentation = Presentation(pptx_file)
print(f"Loaded presentation with {len(presentation.slides)} slides")

# Narration notes
narration_notes = {
    "Code Snippet Maintenance in Documentation": "Welcome. This short video introduces the workflow for code maintenance in documentation, specifically what we do for Azure Machine Learning and Azure AI Foundry.",
    "Definitions": "First a few definitions. There are three different types of repositories involved in this process. To try to keep them straight, we'll call them the Code repo, the Docs repo, and the Maintenance repo.",
    "Learn site": "Here's an example article on the learn site, which contains some code.",
    "GitHub": "Here's the actual code content, stored in a Code Repo. Note the comments that start and stop the block of code called chat_completion.",
    "How to reference": "To use code from an external Code Repo, set a path_to_root entry in the config file. I use the repo name followed by the branch name.",
    "Article markdown": "Now that you have a path_to_root, use it to reference the file. Use id to get to the particular block you want to display.",
    "Advantages": "There are many advantages. The code is authored in an editor that can show typos or mistakes. It's runnable. Styling can be enforced. The repo is maintained by teams using the code. It can be set up with automated testing. And it provides a single source of truth.",
    "PROBLEMS": "Problems include file deletion, renaming, or content changes that break references. Also, updates to code files aren't reflected in docs until rebuilt.",
    "SOLUTION: Monitor the Code Repo": "We monitor the code repo using CODEOWNERS to protect referenced files. Any changes require our review.",
    "Maintenance Process": "Our process includes daily and weekly tasks using Python scripts in the Maintenance Repo. These run in Codespaces and rotate monthly among team members.",
    "Daily: 1. run find-prs": "We check for PRs needing review across multiple repos using a script that outputs a markdown report.",
    "2. run pr-report": "Open the markdown file and view the table. Run the code to check for issues before approving PRs.",
    "3. approve PR": "Approve the PR if no issues. If problems exist, fix the docs first. Consider release branches during events like Build or Ignite.",
    "Weekly: 1. find-snippets": "Search docs for changes in referenced files. This builds CODEOWNERS files and a CSV for lookup.",
    "2. Update CODEOWNERS": "If changes exist, copy the new content and update the associated CODEOWNERS file.",
    "2b. Edit and replace lines": "Edit the file, replace outdated lines, and create a PR to update it.",
    "3. Update Docs": "Run merge-report to find impacted docs. Update metadata to trigger rebuilds.",
    "Update maintenance repo": "Update the Maintenance Repo files so the next person has the latest version.",
    "Questions?": "If you have any questions, feel free to reach out.",
    "Resources": "More info at github.com/sdgilley/content-maintenance and the Microsoft Learn platform manual."
}

# Try to export slides automatically
print("\nAttempting to export slides as images...")
success = False

# Try PowerShell method first (best for Windows)
print("Trying PowerShell export...")
success = export_slides_as_images_powershell(pptx_file)

# If PowerShell failed, try LibreOffice
if not success:
    print("Trying LibreOffice export...")
    success = export_slides_as_images_libreoffice(pptx_file)

# Look for existing exported slides
slides_dir, slide_files = find_exported_slides()

if not slides_dir:
    print("Could not automatically export slides.")
    export_slides_manual_instruction(pptx_file)
    print("\nAfter manually exporting, run this script again.")
    exit()

print(f"Found {len(slide_files)} exported slide images in {slides_dir}/")
for i, filename in enumerate(slide_files[:5]):  # Show first 5
    print(f"  {i+1}: {filename}")
if len(slide_files) > 5:
    print(f"  ... and {len(slide_files)-5} more")

# Create video from exported slides
video_clips = []
print("\nCreating video from slide images...")

for i, slide_file in enumerate(slide_files):
    slide_path = os.path.join(slides_dir, slide_file)
    print(f"\n--- Processing slide {i+1}/{len(slide_files)}: {slide_file} ---")
    
    try:
        # Load the slide image
        slide_img = Image.open(slide_path)
        print(f"Loaded slide image: {slide_img.size}")
        
        # Resize if needed to standard video dimensions
        # Keep aspect ratio but fit within 1280x720
        slide_img.thumbnail((1280, 720), Image.Resampling.LANCZOS)
        
        # Create a 1280x720 canvas and center the slide
        canvas = Image.new("RGB", (1280, 720), color="white")
        x_offset = (1280 - slide_img.width) // 2
        y_offset = (720 - slide_img.height) // 2
        canvas.paste(slide_img, (x_offset, y_offset))
        
        # Save the processed slide
        processed_path = f"slide_images/processed_slide_{i+1}.png"
        os.makedirs("slide_images", exist_ok=True)
        canvas.save(processed_path)
        
        # Create video clip
        duration = 4  # 4 seconds per slide
        image_clip = ImageClip(processed_path).with_duration(duration)
        video_clips.append(image_clip)
        print(f"Created video clip {i+1} with duration {duration}s")
        
    except Exception as e:
        print(f"Error processing slide {i+1}: {e}")

# Combine clips into final video
if video_clips:
    print(f"\nCombining {len(video_clips)} video clips into final video...")
    final_video = concatenate_videoclips(video_clips, method="compose")
    print("Video clips concatenated successfully")

    output_filename = "code_maintenance_process_SLIDES.mp4"
    print(f"Writing final video file: {output_filename}")
    final_video.write_videofile(output_filename, fps=24)
    print("Final video file written successfully!")

    # Cleanup
    print("Starting cleanup...")
    for clip in video_clips:
        clip.close()
    final_video.close()
    
    print(f"\nVideo created successfully: {output_filename}")
    print(f"Duration: {len(video_clips) * 4} seconds ({len(video_clips)} slides x 4 seconds each)")
else:
    print("No video clips were created!")
