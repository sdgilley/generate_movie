# GitHub Copilot Instructions for Generate Movie Project

## Project Overview
This is a PowerPoint-to-video conversion tool that uses Azure AI Speech Services to generate professional videos with AI narration. The project converts PowerPoint presentations into high-quality videos by:

1. Extracting slides as high-resolution images
2. Reading narration text from PowerPoint slide notes
3. Generating natural speech audio using Azure AI Speech
4. Combining slides and audio into a final video

## Key Technologies
- **Python**: Main programming language
- **Azure AI Speech Services**: For AI narration generation
- **MoviePy**: Video processing and assembly
- **LibreOffice + ImageMagick**: Cross-platform slide export (macOS/Linux)
- **PowerShell**: Windows slide export
- **GitHub Codespaces**: Cloud development environment

## Code Style & Standards
- Use descriptive variable names and comprehensive error handling
- Include progress indicators for long-running operations
- Maintain cross-platform compatibility (Windows/macOS/Linux)
- Follow fallback patterns: try multiple methods, gracefully degrade
- Clean up temporary files after processing
- Use environment variables for configuration
- In markdown use 1. for ordered lists and - for unordered lists

## Architecture Patterns
- **Modular design**: Separate utilities for different functions
- **Fallback hierarchy**: Multiple export methods with graceful degradation
- **Dynamic filename generation**: Auto-generate output names from input files
- **Comprehensive cleanup**: Remove intermediate files after processing
- **Cross-platform support**: Handle OS-specific dependencies and commands

## File Organization
- `utilities/`: Core functionality modules
- `exported_slides/`: Generated slide images
- `audio_clips/`: Generated narration audio files
- `slide_images/`: Temporary processed images (cleaned up)
- `.env`: Configuration file (not committed)
- `.devcontainer/`: GitHub Codespaces configuration

## Key Utilities
- `utilities/generate_from_slides.py`: Multi-platform slide export with fallback methods
- `utilities/generate_audio.py`: Azure Speech Services integration
- `utilities/generate_with_azure_audio.py`: Main orchestration script
- `utilities/filename_utils.py`: Dynamic filename generation
- `utilities/cleanup.py`: Temporary file management

## Development Guidelines
1. **Error Handling**: Always include try-catch blocks with informative error messages
2. **Progress Feedback**: Show progress for operations that take time
3. **Platform Detection**: Use `platform.system()` to handle OS differences
4. **Dependency Checks**: Verify required tools are installed before using them
5. **Cleanup**: Always clean up temporary files, even on errors
6. **Configuration**: Use `.env` file for user settings, avoid hardcoded values
7. **Logging**: Provide clear, actionable feedback to users

## Testing Approach
- Test on multiple platforms (Windows, macOS, Linux)
- Verify slide export quality and visual fidelity
- Test with various PowerPoint file formats and sizes
- Validate audio generation quality and timing
- Confirm cleanup operations work correctly

## Common Patterns
When working with this codebase:
- Always check for existing files before creating new ones
- Use absolute paths for file operations
- Implement multiple fallback methods for critical operations
- Provide clear user feedback for each processing step
- Handle edge cases gracefully (empty slides, missing notes, etc.)

## Azure Integration
- Use environment variables for Azure credentials
- Implement proper error handling for API calls
- Support multiple voice options and settings
- Handle rate limiting and quota issues gracefully

## Codespaces Considerations
- Some operations require manual steps due to Linux limitations
- Provide alternative workflows for cloud environments
- Include clear upload/download instructions for users
- Handle missing GUI applications gracefully
