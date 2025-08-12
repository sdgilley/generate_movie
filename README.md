# Generate Movie from PowerPoint

Automatically convert PowerPoint presentations into professional videos with high-quality AI narration using Azure Speech Services.

![Process Diagram](./media/diagram.png)

## 🚀 Quick Start


### Option 1: GitHub Codespaces

[![Open in GitHub Codespaces](https://github.com/codespaces/badge.svg)](https://codespaces.new/sdgilley/generate_movie)

#### ⚠️ Important: Manual Slide Export Required in Codespaces

Due to Linux limitations, automatic slide export may only capture text. For full slide visuals, follow these steps:

1. **Export your slides as PNG images from PowerPoint:**
   - Open your presentation in PowerPoint (Windows or macOS)
   - Go to `File > Export > Change File Type > PNG Portable Network Graphics Format`
   - Click `Save As`, choose a folder (e.g., `exported_slides`)
   - When prompted, select `All Slides`
   - This will create individual PNG files for each slide
1. **Export your slides as PNG images from PowerPoint:**
. **Upload all PNG files to the `exported_slides/` folder in your Codespace**
1. **Export your slides as PNG images from PowerPoint:**
. **Upload your original PowerPoint file (`.pptx`) to the workspace**
1. **Export your slides as PNG images from PowerPoint:**
. **Configure your Azure credentials** in the `.env` file:

   ```env
   SPEECH_KEY=your_azure_speech_key_here
   ENDPOINT=https://your-region.api.cognitive.microsoft.com
   POWERPOINT_FILE=your_presentation.pptx
   ```

1. **Export your slides as PNG images from PowerPoint:**
. **Upload your PowerPoint file** to the codespace
1. **Export your slides as PNG images from PowerPoint:**
1. **Generate your video**: Use Ctrl+Shift+P → "Tasks: Run Task" → "Convert PowerPoint to Video"
1. Delete the files in `uploaded_slides/` when you're done so they won't be used for a new project in the codespace.

**Note:** The code will automatically use the PNGs in `uploaded_slides/` if present and will not attempt to generate slides from text. If no PNGs are found, it will fall back to text-only images.

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

1. **Install additional system dependencies (macOS only):**

   ```bash
   # Install LibreOffice for PowerPoint conversion
   brew install --cask libreoffice
   
   # Install ImageMagick for image processing
   brew install imagemagick
   
   # Install Ghostscript (required for ImageMagick PDF processing)
   brew install ghostscript
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
