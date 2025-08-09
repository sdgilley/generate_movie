#!/usr/bin/env python3
"""
Utility functions for filename handling
"""

import os
from pathlib import Path

def generate_output_filename(pptx_file, suffix=""):
    """
    Generate output video filename based on PowerPoint filename
    
    Args:
        pptx_file (str): Path to the PowerPoint file
        suffix (str): Suffix to add before the extension
    
    Returns:
        str: Generated video filename
    
    Examples:
        "test-ppt.pptx" -> "test-ppt_WITH_AZURE_AUDIO.mp4"
        "content_maintenance_process.pptx" -> "content_maintenance_process_WITH_AZURE_AUDIO.mp4"
        "presentation.ppt" -> "presentation_WITH_AZURE_AUDIO.mp4"
    """
    # Get the file path object
    pptx_path = Path(pptx_file)
    
    # Get the stem (filename without extension)
    base_name = pptx_path.stem
    
    # Create the new filename with suffix and .mp4 extension
    output_filename = f"{base_name}{suffix}.mp4"
    
    return output_filename

def get_powerpoint_file():
    """Get the PowerPoint file from environment or default"""
    return os.environ.get('POWERPOINT_FILE', 'content_maintenance_process.pptx')

def get_output_video_name():
    """Generate output video name based on PowerPoint file"""
    pptx_file = get_powerpoint_file()
    return generate_output_filename(pptx_file)

if __name__ == "__main__":
    # Test the function
    test_files = [
        "test-ppt.pptx",
        "content_maintenance_process.pptx", 
        "my_presentation.ppt",
        "slides.pptx"
    ]
    
    print("Testing filename generation:")
    for file in test_files:
        output = generate_output_filename(file)
        print(f"  {file} -> {output}")
