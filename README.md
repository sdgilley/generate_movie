# Generate Movie from PowerPoint

Automatically convert PowerPoint presentations into professional videos with high-quality AI narration using Azure Speech Services.

![Process Diagram](./media/diagram.png)

## 🚀 Quick Start

### Option 1: GitHub Codespaces (Recommended)

[![Open in GitHub Codespaces](https://github.com/codespaces/badge.svg)](https://codespaces.new/sdgilley/generate_movie)

1. **Click the "Open in GitHub Codespaces" button above**
2. **Wait for the environment to set up** (automatically installs dependencies)
3. **Configure your Azure credentials** in the `.env` file:

   ```bash
   SPEECH_KEY=your_azure_speech_key_here
   ENDPOINT=https://your-region.api.cognitive.microsoft.com
   POWERPOINT_FILE=your_presentation.pptx
   ```

4. **Upload your PowerPoint file** to the codespace
5. **Generate your video**: Use Ctrl+Shift+P → "Tasks: Run Task" → "Convert PowerPoint to Video"

### Option 2: Local Development

1. **Create and activate a virtual environment:**
   ```bash
   # Create virtual environment
   python -m venv venv
   
   # Activate virtual environment
   # On Windows:
   venv\Scripts\activate
   # On macOS/Linux:
   source venv/bin/activate
   ```

1. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

1. **Set up Azure Speech Services:**
   - Create an Azure Cognitive Services Speech resource
   - Copy .env.example to .env and add your Azure credentials:

   ```env
   AZURE_SPEECH_KEY=your_speech_key_here
   AZURE_SPEECH_REGION=your_region_here
   ```

    - Edit the rest of the .env as needed.

1. **Add narration to your PowerPoint:**
   - Open your PowerPoint file
   - Add narration text to the **Notes** section of each slide

1. **Generate your video:**

   ```bash
   python generate_with_azure_audio.py
   ```

## Project Structure

- generate_with_azure_audio.py - Main script for video generation
- generate_from_slides.py - Export slides from PowerPoint
- requirements.txt - Python dependencies
- content_maintenance_process.pptx - Example PowerPoint file

## Configuration

Edit the .env file to customize:

```env
# PowerPoint file to process
POWERPOINT_FILE=your_presentation.pptx

# Note: Output video filename is automatically generated

# Narration source: 'slide_notes' or 'external_file'
NARRATION_SOURCE=slide_notes

# Pause duration before narration (seconds)
PAUSE_DURATION=1.5
```

## How It Works

1. **Slide Export** - Extracts slides from PowerPoint as high-resolution images
1. **Narration Extraction** - Reads narration text from slide notes
1. **Audio Generation** - Creates natural speech using Azure Speech Services
1. **Video Assembly** - Combines slides with audio using optimal timing:
   - Slide appears immediately
   - reading pause at beginning of each slide (configure in .env file, default 1.5 sec)
   - Narration plays while slide remains visible
   - Smooth transition to next slide

## Requirements

- Python 3.7+
- Azure Cognitive Services Speech subscription
- PowerPoint (for slide export)
- See requirements.txt for Python packages

## Contributing

Feel free to submit issues and enhancement requests!

## License

This project is open source and available under the MIT License.
