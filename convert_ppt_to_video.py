#!/usr/bin/env python3
"""
PowerPoint to Video Converter with Azure Speech Services

This script automates the complete process of converting a PowerPoint presentation
to a video with narration using Azure Speech Services.

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
    
    # Check for PowerPoint file
    pptx_file = "content_maintenance_process.pptx"
    if not os.path.exists(pptx_file):
        print(f"❌ ERROR: PowerPoint file not found: {pptx_file}")
        print("Please ensure your PowerPoint file is in the current directory.")
        return False
    print(f"✅ Found PowerPoint file: {pptx_file}")
    
    # Check for .env file
    env_file = ".env"
    if not os.path.exists(env_file):
        print(f"❌ ERROR: Environment file not found: {env_file}")
        print("Please create a .env file with your Azure Speech Services credentials:")
        print("SPEECH_KEY=your_azure_speech_key")
        print("ENDPOINT=https://your-region.api.cognitive.microsoft.com")
        return False
    print(f"✅ Found environment file: {env_file}")
    
    # Check for required Python files
    required_files = ["generate_from_slides.py", "generate_audio.py", "generate_with_azure_audio.py"]
    for file in required_files:
        if not os.path.exists(file):
            print(f"❌ ERROR: Required Python file not found: {file}")
            return False
        print(f"✅ Found: {file}")
    
    print("✅ All prerequisites met!")
    return True

def run_command(command, description):
    """Run a command and handle errors"""
    print(f"\n🚀 Running: {description}")
    print(f"Command: {command}")
    
    try:
        # Get the Python executable path
        python_exe = sys.executable
        if ".venv" in python_exe:
            # We're in a virtual environment, use it
            cmd = f'"{python_exe}" {command}'
        else:
            # Try to use the virtual environment if it exists
            venv_python = Path("C:/git/movies/.venv/Scripts/python.exe")
            if venv_python.exists():
                cmd = f'"{venv_python}" {command}'
            else:
                cmd = f'python {command}'
        
        result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
        
        if result.returncode == 0:
            print("✅ Success!")
            if result.stdout:
                print("Output:")
                print(result.stdout)
            return True
        else:
            print("❌ Error occurred!")
            if result.stderr:
                print("Error details:")
                print(result.stderr)
            if result.stdout:
                print("Output:")
                print(result.stdout)
            return False
            
    except Exception as e:
        print(f"❌ Exception occurred: {e}")
        return False

def check_output_files():
    """Check if expected output files were created"""
    print_step(4, "Checking output files")
    
    expected_files = [
        "code_maintenance_process_WITH_AZURE_AUDIO.mp4",
        "exported_slides/",
        "audio_clips/",
        "slide_images/"
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
    success = run_command("generate_from_slides.py", "Exporting PowerPoint slides to images")
    if not success:
        print("❌ Failed to export slides. Please check the error messages above.")
        return False
    
    # Step 2: Test Azure Speech Services
    print_step(2, "Testing Azure Speech Services")
    success = run_command("generate_audio.py", "Testing Azure Speech Services connection")
    if not success:
        print("❌ Azure Speech Services test failed. Please check your credentials in .env file.")
        return False
    
    # Step 3: Generate final video with audio
    print_step(3, "Generating final video with Azure Speech narration")
    success = run_command("generate_with_azure_audio.py", "Creating final video with audio and pauses")
    if not success:
        print("❌ Failed to generate final video. Please check the error messages above.")
        return False
    
    # Step 4: Verify outputs
    if not check_output_files():
        print("⚠️  Some expected output files were not found.")
    
    # Calculate total time
    end_time = time.time()
    total_time = end_time - start_time
    
    print_header("Process Complete!")
    print(f"✅ Total processing time: {total_time:.1f} seconds ({total_time/60:.1f} minutes)")
    print(f"🎬 Your video is ready: code_maintenance_process_WITH_AZURE_AUDIO.mp4")
    
    # Final summary
    print("\n📊 Files created:")
    print("  • code_maintenance_process_WITH_AZURE_AUDIO.mp4 - Final video")
    print("  • exported_slides/ - Individual slide images")
    print("  • audio_clips/ - Generated audio files")
    print("  • slide_images/ - Processed slide images")
    print("  • test_audio/ - Audio test files")
    
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
