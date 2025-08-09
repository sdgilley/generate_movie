#!/bin/bash
# Codespace Setup Script for PowerPoint to Video Converter
# This script helps set up the environment in GitHub Codespaces

echo "🚀 Setting up PowerPoint to Video Converter in Codespace..."

# Check if .env file exists
if [ ! -f .env ]; then
    echo "📝 Creating .env file from template..."
    cp .env.example .env
    echo "✅ Created .env file. Please edit it with your Azure Speech Service credentials."
    echo ""
    echo "📋 To get started:"
    echo "1. Edit .env file with your Azure Speech Services credentials"
    echo "2. Add your PowerPoint file to the workspace"
    echo "3. Update POWERPOINT_FILE in .env to point to your file"
    echo "4. Run: python convert_ppt_to_video.py"
else
    echo "✅ .env file already exists"
fi

# Check if requirements are installed
echo "🔍 Checking Python environment..."
if python -c "import azure.cognitiveservices.speech" 2>/dev/null; then
    echo "✅ Azure Speech SDK is installed"
else
    echo "📦 Installing Python dependencies..."
    pip install -r requirements.txt
fi

# Check for sample PowerPoint file
if [ ! -f "test-ppt.pptx" ] && [ ! -f "content_maintenance_process.pptx" ]; then
    echo "⚠️  No sample PowerPoint file found."
    echo "   Please upload your PowerPoint file to the codespace."
fi

echo ""
echo "🎉 Setup complete! Here's what you can do next:"
echo ""
echo "1. 📝 Edit your .env file:"
echo "   code .env"
echo ""
echo "2. 🎯 Test the system (if you have a PowerPoint file):"
echo "   python convert_ppt_to_video.py"
echo ""
echo "3. 📚 Read the documentation:"
echo "   code README.md"
echo ""
echo "4. 🔧 Test individual components:"
echo "   python utilities/filename_utils.py"
echo ""
echo "Happy video creation! 🎬"
