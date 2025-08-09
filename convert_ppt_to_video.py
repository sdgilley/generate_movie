#!/usr/bin/env python3
"""
PowerPoint to Video Converter with Azure Speech Services

This script automates the complete process of converting a PowerPoint presentation
to a video with narration using Azure Speech Services.

    print("📊 Files created:")
    print("  • ppt_to_mp4WITH_AZURE_AUDIO.mp4 - Final video")
    print("  • exported_slides/ - Individual slide images")
    print("  • audio_clips/ - Generated audio files")
    print("  • slide_images/ - Processed slide images")
    print("  • test_audio/ - Audio test files")th narration using Azure Speech Services.

Requirements:
- PowerPoint presentation file: content_maintenance_process.pptx
- Azure Speech Services credentials in .env file
- Python packages: see requirements.txt

Process:
1. Export PowerPoint slides as images
2. Generate audio narration using Azure Speech Services
3. Combine slides and audio into final video with pauses
"""

import os
import sys
import subprocess
import time
from pathlib import Path
from dotenv import load_dotenv

# Import utility functions
from utilities.generate_from_slides import main as export_slides
from utilities.generate_audio import test_audio_generation
from utilities.generate_with_azure_audio import main as generate_video_with_audio
from utilities.filename_utils import get_powerpoint_file, get_output_video_name

# Load environment variables
load_dotenv()

def print_header(title):
    """Print a formatted header"""
    print("\n" + "="*60)
    print(f" {title}")
    print("="*60)

def print_step(step_num, description):
    """Print a formatted step"""
    print(f"\n🔹 Step {step_num}: {description}")

def check_prerequisites():
    """Check if all required files and credentials exist"""
    print_step(0, "Checking prerequisites")
    
    # Check for .env file
    env_file = ".env"
    if not os.path.exists(env_file):
        print(f"❌ ERROR: Environment file not found: {env_file}")
        print("Please create a .env file with your Azure Speech Services credentials:")
        print("SPEECH_KEY=your_azure_speech_key")
        print("ENDPOINT=https://your-region.api.cognitive.microsoft.com")
        return False
    print(f"✅ Found environment file: {env_file}")

    # Check for PowerPoint file
    pptx_file = get_powerpoint_file()
    if not os.path.exists(pptx_file):
        print(f"❌ ERROR: PowerPoint file not found: {pptx_file}")
        print("Please ensure your PowerPoint file is in the current directory.")
        return False
    print(f"✅ Found PowerPoint file: {pptx_file}")
    
    # Check for required Python files
    required_files = ["utilities/generate_from_slides.py", "utilities/generate_audio.py", "utilities/generate_with_azure_audio.py"]
    for file in required_files:
        if not os.path.exists(file):
            print(f"❌ ERROR: Required Python file not found: {file}")
            return False
        print(f"✅ Found: {file}")
    
    print("✅ All prerequisites met!")
    return True

def check_output_files():
    """Check if expected output files were created"""
    print_step(4, "Checking output files")
    
    # Get filenames from utility functions
    output_video = get_output_video_name()
    
    expected_files = [
        output_video,
        "exported_slides/",
        "audio_clips/"
        # Note: slide_images/ is intentionally excluded as it's cleaned up after processing
    ]
    
    all_found = True
    for file_path in expected_files:
        if os.path.exists(file_path):
            if os.path.isdir(file_path):
                file_count = len([f for f in os.listdir(file_path) if not f.startswith('.')])
                print(f"✅ Found directory: {file_path} ({file_count} files)")
            else:
                file_size = os.path.getsize(file_path)
                print(f"✅ Found file: {file_path} ({file_size:,} bytes)")
        else:
            print(f"❌ Missing: {file_path}")
            all_found = False
    
    return all_found

def main():
    """Main driver function"""
    print_header("PowerPoint to Video Converter with Azure Speech")
    print("This script will convert your PowerPoint presentation to a video with narration.")
    
    start_time = time.time()
    
    # Step 0: Check prerequisites
    if not check_prerequisites():
        print("\n❌ Prerequisites not met. Please fix the issues above and try again.")
        return False
    
    # Step 1: Export slides from PowerPoint
    print_step(1, "Exporting slides from PowerPoint as images")
    try:
        success = export_slides()
        if not success:
            print("❌ Failed to export slides.")
            return False
        print("✅ Slides exported successfully!")
    except Exception as e:
        print(f"❌ Error exporting slides: {e}")
        return False
    
    # Step 2: Test Azure Speech Services
    print_step(2, "Testing Azure Speech Services")
    try:
        success = test_audio_generation()
        if not success:
            print("❌ Azure Speech Services test failed. Please check your credentials in .env file.")
            return False
        print("✅ Azure Speech Services test passed!")
    except Exception as e:
        print(f"❌ Error testing Azure Speech Services: {e}")
        return False
    
    # Step 3: Generate final video with audio
    print_step(3, "Generating final video with Azure Speech narration")
    try:
        success = generate_video_with_audio()
        if not success:
            print("❌ Failed to generate final video.")
            return False
        print("✅ Final video generated successfully!")
    except Exception as e:
        print(f"❌ Error generating final video: {e}")
        return False
    
    # Step 4: Verify outputs
    if not check_output_files():
        print("⚠️  Some expected output files were not found.")
    
    # Calculate total time
    end_time = time.time()
    total_time = end_time - start_time
    
    # Get output video name from utility function
    output_video = get_output_video_name()
    
    print_header("Process Complete!")
    print(f"✅ Total processing time: {total_time:.1f} seconds ({total_time/60:.1f} minutes)")
    print(f"🎬 Your video is ready: {output_video}")
    
    # Final summary
    print("\n📊 Files created:")
    print(f"  • {output_video} - Final video")
    print("  • exported_slides/ - Individual slide images")
    print("  • audio_clips/ - Generated audio files")
    print("\n🧹 Cleaned up:")
    print("  • slide_images/ - Temporary processed slide images")
    print("  • test_audio/ - Temporary audio test files")
    
    print("\n🎉 Success! Your PowerPoint presentation has been converted to video with narration.")
    return True

if __name__ == "__main__":
    try:
        success = main()
        exit_code = 0 if success else 1
        sys.exit(exit_code)
    except KeyboardInterrupt:
        print("\n\n⚠️  Process interrupted by user (Ctrl+C)")
        sys.exit(1)
    except Exception as e:
        print(f"\n\n❌ Unexpected error: {e}")
        sys.exit(1)
