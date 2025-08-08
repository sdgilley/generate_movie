import os
import azure.cognitiveservices.speech as speechsdk
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

def generate_audio_file(text, output_path, voice_name=None):
    """
    Generate audio file from text using Azure Speech Services
    
    Args:
        text (str): Text to convert to speech
        output_path (str): Path where to save the audio file (should end with .wav)
        voice_name (str): Voice to use for synthesis (if None, uses .env setting)
    
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        # Get Azure Speech credentials
        speech_key = os.environ.get('SPEECH_KEY')
        speech_region = "eastus2"  # Your region
        
        # Get voice name from .env if not provided
        if voice_name is None:
            voice_name = os.environ.get('VOICE_NAME', 'en-US-AvaMultilingualNeural')
        
        if not speech_key:
            print("Error: SPEECH_KEY not found in environment variables")
            return False
        
        # Create speech config
        speech_config = speechsdk.SpeechConfig(subscription=speech_key, region=speech_region)
        speech_config.speech_synthesis_voice_name = voice_name
        
        # Configure audio output to file
        audio_config = speechsdk.audio.AudioOutputConfig(filename=output_path)
        
        # Create synthesizer
        speech_synthesizer = speechsdk.SpeechSynthesizer(speech_config=speech_config, audio_config=audio_config)
        
        print(f"Generating audio for: {text[:50]}...")
        
        # Synthesize speech
        speech_synthesis_result = speech_synthesizer.speak_text_async(text).get()
        
        if speech_synthesis_result.reason == speechsdk.ResultReason.SynthesizingAudioCompleted:
            print(f"Audio saved successfully: {output_path}")
            return True
        elif speech_synthesis_result.reason == speechsdk.ResultReason.Canceled:
            cancellation_details = speech_synthesis_result.cancellation_details
            print(f"Speech synthesis canceled: {cancellation_details.reason}")
            if cancellation_details.reason == speechsdk.CancellationReason.Error:
                if cancellation_details.error_details:
                    print(f"Error details: {cancellation_details.error_details}")
            return False
        else:
            print(f"Unexpected result: {speech_synthesis_result.reason}")
            return False
            
    except Exception as e:
        print(f"Error generating audio: {e}")
        return False

def test_audio_generation():
    """Test the audio generation function"""
    print("Testing Azure Speech Services...")
    
    # Create test directory
    os.makedirs("test_audio", exist_ok=True)
    
    test_text = "Hello, this is a test of the Azure Speech Services text to speech functionality."
    output_file = "test_audio/test_azure_speech.wav"
    
    success = generate_audio_file(test_text, output_file)
    
    if success:
        print(f"Test successful! Audio file created: {output_file}")
        # Check file size
        if os.path.exists(output_file):
            file_size = os.path.getsize(output_file)
            print(f"File size: {file_size} bytes")
        return True
    else:
        print("Test failed!")
        return False

if __name__ == "__main__":
    # Test the function when script is run directly
    test_audio_generation()