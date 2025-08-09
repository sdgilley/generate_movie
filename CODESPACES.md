# Using PowerPoint to Video Converter in GitHub Codespaces

This guide explains how to use the PowerPoint to Video Converter in GitHub Codespaces.

## Getting Started

### 1. Open in Codespaces

Click the "Open in GitHub Codespaces" button in the README, or:
1. Go to the repository on GitHub
2. Click the green "Code" button
3. Select "Codespaces" tab
4. Click "Create codespace on main"

### 2. Wait for Setup

The codespace will automatically:
- Install Python dependencies
- Set up the development environment
- Run the setup script
- Create a `.env` file from the template

### 3. Configure Azure Speech Services

Edit the `.env` file with your Azure credentials:
```bash
code .env
```

Add your Azure Speech Services credentials:
```env
SPEECH_KEY=your_azure_speech_key_here
ENDPOINT=https://your-region.api.cognitive.microsoft.com
```

### 4. Upload Your PowerPoint File

1. Drag and drop your PowerPoint file into the file explorer
2. Update the `POWERPOINT_FILE` in `.env` to match your filename

### 5. Generate Your Video

#### Option A: Using VS Code Tasks (Recommended)
1. Press `Ctrl+Shift+P` (or `Cmd+Shift+P` on Mac)
2. Type "Tasks: Run Task"
3. Select "Convert PowerPoint to Video"

#### Option B: Using Terminal
```bash
python convert_ppt_to_video.py
```

## Available VS Code Tasks

Access via `Ctrl+Shift+P` → "Tasks: Run Task":

- **Convert PowerPoint to Video** - Full conversion process
- **Test Slide Export Only** - Test slide extraction
- **Test Azure Speech Services** - Test audio generation
- **Test Filename Generation** - Test filename utilities
- **Setup Environment** - Re-run setup script

## Debugging

You can debug the application using VS Code's built-in debugger:

1. Set breakpoints in the code
2. Press `F5` or go to Run and Debug view
3. Select "Debug PowerPoint Converter"
4. The debugger will start and pause at your breakpoints

## File Structure in Codespaces

```
/workspaces/generate_movie/
├── .devcontainer/          # Codespace configuration
├── .vscode/               # VS Code settings and tasks
├── .github/               # GitHub workflows
├── utilities/             # Core processing modules
├── media/                 # Assets and output images
├── .env                   # Your configuration (created from .env.example)
├── convert_ppt_to_video.py # Main script
└── README.md              # Documentation
```

## Tips for Codespaces

1. **File Uploads**: Drag and drop files directly into the file explorer
2. **Terminal**: Use the integrated terminal for running commands
3. **Extensions**: Python extensions are pre-installed and configured
4. **Environment**: Everything is pre-configured - just add your Azure credentials!
5. **Persistence**: Your codespace saves your work automatically

## Troubleshooting

### Azure Speech Services Issues
- Verify your `SPEECH_KEY` and `ENDPOINT` in `.env`
- Test with: "Tasks: Run Task" → "Test Azure Speech Services"

### PowerPoint File Issues
- Ensure your file is uploaded to the codespace
- Check the `POWERPOINT_FILE` setting in `.env`
- Supported formats: `.pptx`, `.ppt`

### Permission Issues
- The setup script should handle permissions automatically
- If needed, run: `chmod +x setup_codespace.sh`

## Getting Help

If you encounter issues:
1. Check the terminal output for error messages
2. Run individual test tasks to isolate problems
3. Verify your `.env` configuration
4. Check that your PowerPoint file is properly uploaded

Happy video creation! 🎬
