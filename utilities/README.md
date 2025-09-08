# Utilities Directory

This directory contains the core modules for PowerPoint-to-video conversion:

## Core Modules

### `generate_audio.py`
Core text-to-speech utility module that handles Azure Speech Services integration:
- Provides the `generate_audio_file()` function for text-to-speech conversion
- Manages Azure Speech Services authentication (subscription key or token)
- Handles audio file generation and error handling
- Not meant to be run directly; used as a library by other modules

### `generate_with_azure_audio.py`
Main script that orchestrates the PowerPoint-to-video conversion process:
- This is the script you should run to create videos
- Uses `generate_audio.py` for text-to-speech functionality
- Handles slide extraction, audio generation, and video assembly
- Run with: `python generate_with_azure_audio.py`

## Usage

1. Set up your `.env` file with Azure credentials
2. Run `generate_with_azure_audio.py` to create a video from your PowerPoint
3. The script will:
   - Extract slides as images
   - Convert notes to speech using `generate_audio.py`
   - Combine everything into a final video

## File Organization
```
utilities/
├── generate_audio.py        # Core text-to-speech library
├── generate_with_azure_audio.py   # Main script to run
└── README.md               # This file
```
