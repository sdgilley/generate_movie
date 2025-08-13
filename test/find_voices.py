#!/usr/bin/env python3
"""
Azure Speech Voices List

This script displays a clean list of all available voice names in your Azure Speech region.
"""

import os
import sys
from dotenv import load_dotenv
import azure.cognitiveservices.speech as speechsdk

def list_available_voices(speech_key, region):
    """Get and display list of available voice names"""
    try:
        # Create speech config
        speech_config = speechsdk.SpeechConfig(subscription=speech_key, region=region)
        
        # Create synthesizer
        synthesizer = speechsdk.SpeechSynthesizer(speech_config=speech_config, audio_config=None)
        
        # Get voices
        print(f"Fetching available voices for region '{region}'...")
        result = synthesizer.get_voices_async().get()
        
        if result.reason == speechsdk.ResultReason.VoicesListRetrieved:
            voices = result.voices
            
            # Group voices by locale
            voices_by_locale = {}
            for voice in voices:
                locale = voice.locale
                if locale not in voices_by_locale:
                    voices_by_locale[locale] = []
                voices_by_locale[locale].append(voice.name)
            
            # Display English voices only
            en_locales = [locale for locale in voices_by_locale.keys() if locale.startswith('en-')]
            
            print(f"\n" + "=" * 60)
            print(f" ENGLISH VOICES ({len([v for locale in en_locales for v in voices_by_locale[locale]])} total)")
            print("=" * 60)
            
            for locale in sorted(en_locales):
                print(f"\n{locale} ({len(voices_by_locale[locale])} voices):")
                for voice_name in sorted(voices_by_locale[locale]):
                    # Clean up the voice name for display
                    clean_name = voice_name.replace("Microsoft Server Speech Text to Speech Voice (", "").replace(")", "")
                    if ", " in clean_name:
                        parts = clean_name.split(", ")
                        if len(parts) >= 2:
                            clean_name = parts[1]  # Just the voice name part
                    print(f"  {clean_name}")
            
            print(f"\n" + "=" * 60)
            print(f" USAGE EXAMPLES")
            print("=" * 60)
            print("Copy any voice name above and use it in your .env file:")
            print("  VOICE_NAME=en-US-JennyNeural")
            print("  VOICE_NAME=en-US-AriaNeural")
            print("  VOICE_NAME=en-GB-SoniaNeural")
            print("\nRecommended voices for different styles:")
            print("  • Conversational: en-US-JennyNeural, en-US-AriaNeural")
            print("  • Professional: en-US-DavisNeural, en-US-JaneNeural")
            print("  • Friendly: en-US-AshleyNeural, en-US-BrandonNeural")
            print("  • British accent: en-GB-SoniaNeural, en-GB-RyanNeural")
            print("  • Australian accent: en-AU-NatashaNeural, en-AU-WilliamNeural")
            
            return True
            
        else:
            print(f"❌ Failed to retrieve voices: {result.reason}")
            return False
            
    except Exception as e:
        print(f"❌ Error fetching voices: {str(e)}")
        return False

def main():
    """Main function to list voices"""
    print("=" * 60)
    print(" Azure Speech Services - Available Voices")
    print("=" * 60)
    
    # Load environment
    load_dotenv()
    
    speech_key = os.getenv('SPEECH_KEY')
    endpoint = os.getenv('ENDPOINT', '')
    
    # Extract region
    region = "eastus"
    if "eastus2" in endpoint:
        region = "eastus2"
    elif "eastus" in endpoint:
        region = "eastus"
    elif "westus2" in endpoint:
        region = "westus2"
    elif "westus" in endpoint:
        region = "westus"
    
    print(f"Configuration:")
    print(f"  Speech Key: {speech_key[:10]}..." if speech_key else "  Speech Key: Not found")
    print(f"  Region: {region}")
    
    if not speech_key:
        print("❌ SPEECH_KEY not found in .env file")
        return
    
    # List available voices
    success = list_available_voices(speech_key, region)
    
    if not success:
        print("\n❌ Failed to retrieve voice list")
        print("💡 Check your Azure Speech Services configuration")

if __name__ == "__main__":
    main()
