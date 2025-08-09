# GitHub Codespaces Configuration

This project is now fully configured to work seamlessly in GitHub Codespaces! Here's what has been added:

## Files Added for Codespace Support

### 🐳 Development Container Configuration
- **`.devcontainer/devcontainer.json`** - Defines the codespace environment
  - Based on Python 3.12 container
  - Pre-installs Python extensions and debugger
  - Automatically runs setup script on creation

### 🔧 VS Code Configuration
- **`.vscode/tasks.json`** - Predefined tasks for common operations
  - Convert PowerPoint to Video (main task)
  - Test individual components
  - Setup environment
- **`.vscode/launch.json`** - Debug configurations
  - Debug main converter
  - Debug slide export
  - Debug audio generation

### 🚀 Setup and Documentation
- **`setup_codespace.sh`** - Automated setup script
  - Creates `.env` from template
  - Installs dependencies
  - Provides helpful guidance
- **`CODESPACES.md`** - Comprehensive usage guide
- **`.env.example`** - Template for configuration

### 🧪 Testing
- **`.github/workflows/test.yml`** - GitHub Actions workflow
  - Tests environment setup
  - Validates imports and dependencies
  - Ensures cross-platform compatibility

## How It Works

1. **One-Click Setup**: Click "Open in GitHub Codespaces" 
2. **Automatic Environment**: Container builds with Python 3.12 and all dependencies
3. **Pre-configured VS Code**: Extensions, tasks, and debug configs ready to use
4. **Guided Setup**: Setup script creates `.env` and provides next steps

## Key Features

✅ **Zero local setup required** - Everything runs in the cloud  
✅ **Cross-platform compatibility** - Works on any device with a browser  
✅ **Pre-configured development environment** - Python, extensions, and tools ready  
✅ **One-click task execution** - Use VS Code tasks for common operations  
✅ **Debugging support** - Full debugging capabilities with breakpoints  
✅ **Automatic dependency management** - All Python packages installed automatically  

## Usage in Codespaces

### Quick Start
1. Click the "Open in GitHub Codespaces" badge in README
2. Wait for environment setup (1-2 minutes)
3. Edit `.env` with your Azure credentials
4. Upload your PowerPoint file
5. Use Ctrl+Shift+P → "Tasks: Run Task" → "Convert PowerPoint to Video"

### Available Tasks
- **Convert PowerPoint to Video** - Full end-to-end conversion
- **Test Slide Export Only** - Test slide image generation
- **Test Azure Speech Services** - Verify audio generation
- **Test Filename Generation** - Check filename utilities
- **Setup Environment** - Re-run setup if needed

### File Management
- Drag & drop PowerPoint files into the file explorer
- Generated videos appear in the workspace root
- All output files are preserved in the codespace

## Benefits

- **No installation required** - Run anywhere with internet connection
- **Consistent environment** - Same setup for all users
- **Easy collaboration** - Share codespace links with team members
- **Resource efficient** - Use cloud compute instead of local resources
- **Always up-to-date** - Latest project state in every new codespace

This makes the PowerPoint to Video Converter accessible to anyone, regardless of their local development setup! 🌟
