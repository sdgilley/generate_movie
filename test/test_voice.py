#!/usr/bin/env python3
"""
Simple Azure Speech Test - Using Working Function

This script uses the same function that works in your video generation
to test Azure Speech Services with better error reporting.
"""

import os
import sys
from dotenv import load_dotenv

def main():
    """Test Azure Speech using the working function"""
    print("=" * 60)
    print(" Azure Speech Test - Using Working Function")
    print("=" * 60)
    
    # Load environment variables
    print("📄 Loading .env file...")
    load_dotenv()
    
    # Show configuration
    speech_key = os.getenv('SPEECH_KEY')
    voice_name = os.getenv('VOICE_NAME')
    
    print(f"✅ SPEECH_KEY: {speech_key[:10]}..." if speech_key else "❌ SPEECH_KEY: Not found")
    print(f"✅ VOICE_NAME: {voice_name}" if voice_name else "ℹ️  VOICE_NAME: Will use default")
    
    # Test with direct audio output (no file)
    try:
        import azure.cognitiveservices.speech as speechsdk

        test_text = f"Hello! This is a test of Azure Speech Services. I'm {voice_name.replace('en-US-', '')}. If you can hear this, everything is working perfectly!"

        print(f"\n🎤 Testing with: '{test_text}'")
        print("🔄 Generating and playing audio...")
        
        # Create speech configuration (same as working script)
        speech_config = speechsdk.SpeechConfig(subscription=speech_key, region="eastus2")
        
        # Create synthesizer with default audio output (speakers)
        speech_synthesizer = speechsdk.SpeechSynthesizer(speech_config=speech_config)
        
        # Synthesize and play speech
        speech_synthesis_result = speech_synthesizer.speak_text_async(test_text).get()
        
        # Check result
        if speech_synthesis_result.reason == speechsdk.ResultReason.SynthesizingAudioCompleted:
            result = True
        elif speech_synthesis_result.reason == speechsdk.ResultReason.Canceled:
            cancellation_details = speechsdk.CancellationDetails(speech_synthesis_result)
            print(f"❌ Speech synthesis canceled: {cancellation_details.reason}")
            if cancellation_details.reason == speechsdk.CancellationReason.Error:
                print(f"❌ Error details: {cancellation_details.error_details}")
            result = False
        else:
            result = False
        
        
        if result:
            print("✅ SUCCESS! Audio played successfully")
            print("🎉 Test completed successfully!")
            print("💡 Your Azure Speech Services is working correctly.")
        else:
            print("❌ FAILED: Audio generation/playback failed")
            print("💡 This usually indicates authentication or configuration issues.")
            print("🔍 Check the error messages above for specific details.")
            
            print("\n📋 Troubleshooting checklist:")
            print("   1. Verify SPEECH_KEY is correct in .env file")
            print("   2. Check if your Azure subscription is active")
            print("   3. Ensure Speech Services resource is in eastus2 region")
            print("   4. Verify you have Speech Services quota remaining")
            
    except ImportError as e:
        print(f"❌ Import error: {e}")
        print("💡 Make sure utilities module is available")
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
    
    print("\n" + "=" * 60)

if __name__ == "__main__":
    main()
