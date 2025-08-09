#!/usr/bin/env python3
"""
Cleanup Script for PowerPoint to Video Converter

This script removes intermediate files created during the video generation process.
Run this after you're satisfied with your final video to free up disk space.

What gets cleaned up:
- slide_images/ (processed slide images)
- audio_clips/ (generated audio files)
- test_audio/ (test files)
- Optionally: exported_slides/ (original slide exports)
"""

import os
import shutil

def remove_directory(dir_path):
    """Remove a directory and all its contents"""
    if os.path.exists(dir_path):
        try:
            shutil.rmtree(dir_path)
            print(f"✅ Removed directory: {dir_path}")
            return True
        except Exception as e:
            print(f"❌ Error removing {dir_path}: {e}")
            return False
    else:
        print(f"⚪ Directory not found: {dir_path}")
        return True

def get_directory_size(dir_path):
    """Get the total size of a directory in bytes"""
    if not os.path.exists(dir_path):
        return 0
    
    total_size = 0
    for dirpath, dirnames, filenames in os.walk(dir_path):
        for filename in filenames:
            file_path = os.path.join(dirpath, filename)
            try:
                total_size += os.path.getsize(file_path)
            except (OSError, FileNotFoundError):
                pass
    return total_size

def format_size(size_bytes):
    """Format size in bytes to human readable format"""
    if size_bytes == 0:
        return "0 B"
    
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size_bytes < 1024.0:
            return f"{size_bytes:.1f} {unit}"
        size_bytes /= 1024.0
    return f"{size_bytes:.1f} TB"

def main():
    print("🧹 PowerPoint to Video Cleanup Tool")
    print("=" * 50)
    
    # Calculate sizes before cleanup
    dirs_to_clean = ["slide_images", "audio_clips", "test_audio", "exported_slides"]
    total_size_before = 0
    
    print("\n📊 Current disk usage:")
    for dir_name in dirs_to_clean:
        size = get_directory_size(dir_name)
        total_size_before += size
        if size > 0:
            print(f"  {dir_name}/: {format_size(size)}")
    
    if total_size_before == 0:
        print("  No cleanup needed - no intermediate files found.")
        return
    
    print(f"\n📦 Total size to clean: {format_size(total_size_before)}")
    
    # Ask user what to clean
    print("\n🗂️  What would you like to clean up?")
    print("1. Essential cleanup (slide_images, test_audio)")
    print("2. Standard cleanup (slide_images, test_audio, audio_clips)")
    print("3. Full cleanup (everything including exported_slides)")
    print("4. Custom selection")
    print("5. Cancel")
    
    while True:
        choice = input("\nEnter your choice (1-5): ").strip()
        
        if choice == "1":
            # Essential cleanup
            dirs_to_remove = ["slide_images", "test_audio"]
            break
        elif choice == "2":
            # Standard cleanup
            dirs_to_remove = ["slide_images", "test_audio", "audio_clips"]
            break
        elif choice == "3":
            # Full cleanup
            dirs_to_remove = ["slide_images", "test_audio", "audio_clips", "exported_slides"]
            break
        elif choice == "4":
            # Custom selection
            dirs_to_remove = []
            for dir_name in dirs_to_clean:
                if os.path.exists(dir_name):
                    response = input(f"Remove {dir_name}/? (y/n): ").strip().lower()
                    if response in ['y', 'yes']:
                        dirs_to_remove.append(dir_name)
            break
        elif choice == "5":
            print("Cleanup cancelled.")
            return
        else:
            print("Invalid choice. Please enter 1-5.")
    
    if not dirs_to_remove:
        print("No directories selected for cleanup.")
        return
    
    # Confirm cleanup
    print(f"\n⚠️  About to remove:")
    for dir_name in dirs_to_remove:
        if os.path.exists(dir_name):
            size = get_directory_size(dir_name)
            print(f"  • {dir_name}/ ({format_size(size)})")
    
    confirm = input("\nProceed with cleanup? (y/n): ").strip().lower()
    if confirm not in ['y', 'yes']:
        print("Cleanup cancelled.")
        return
    
    # Perform cleanup
    print("\n🧹 Cleaning up...")
    success_count = 0
    total_removed = 0
    
    for dir_name in dirs_to_remove:
        size_before = get_directory_size(dir_name)
        if remove_directory(dir_name):
            success_count += 1
            total_removed += size_before
    
    print(f"\n✅ Cleanup completed!")
    print(f"📊 Freed up: {format_size(total_removed)}")
    print(f"🗂️  Removed {success_count}/{len(dirs_to_remove)} directories")
    
    # Show what's left
    print("\n📁 Remaining files:")
    remaining_files = [
        "content_maintenance_process.pptx",
        "code_maintenance_process_WITH_AZURE_AUDIO.mp4",
        ".env"
    ]
    
    for file_name in remaining_files:
        if os.path.exists(file_name):
            if os.path.isfile(file_name):
                size = os.path.getsize(file_name)
                print(f"  • {file_name} ({format_size(size)})")
            else:
                print(f"  • {file_name}/")
    
    print(f"\n🎬 Your video is ready: code_maintenance_process_WITH_AZURE_AUDIO.mp4")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nCleanup cancelled by user.")
    except Exception as e:
        print(f"\nError during cleanup: {e}")
