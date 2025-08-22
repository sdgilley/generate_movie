import os
import types
import tempfile
import generate_audio


def test_generate_audio_uses_access_token(monkeypatch, tmp_path):
    # Ensure subscription key is not set and token is set
    os.environ.pop('SPEECH_KEY', None)
    os.environ['AZURE_ACCESS_TOKEN'] = 'fake-token-123'

    tmp_file = tmp_path / "out.wav"

    class DummyResultReason:
        SynthesizingAudioCompleted = 0
        Canceled = 1

    class DummyResult:
        def __init__(self, reason):
            self.reason = reason
            self.cancellation_details = types.SimpleNamespace(reason=None, error_details=None)

    class DummyFuture:
        def __init__(self, res):
            self._res = res

        def get(self):
            return self._res

    class DummySpeechConfig:
        last = None

        def __init__(self, subscription=None, region=None):
            DummySpeechConfig.last = self
            self.subscription = subscription
            self.region = region
            self.speech_synthesis_voice_name = None
            self.authorization_token = None

    class DummyAudioOutputConfig:
        def __init__(self, filename=None):
            self.filename = filename

    class DummySpeechSynthesizer:
        def __init__(self, speech_config=None, audio_config=None):
            self.speech_config = speech_config
            self.audio_config = audio_config

        def speak_text_async(self, text):
            return DummyFuture(DummyResult(DummyResultReason.SynthesizingAudioCompleted))

    monkeypatch.setattr(generate_audio, 'speechsdk', types.SimpleNamespace(
        SpeechConfig=DummySpeechConfig,
        audio=types.SimpleNamespace(AudioOutputConfig=DummyAudioOutputConfig),
        SpeechSynthesizer=DummySpeechSynthesizer,
        ResultReason=DummyResultReason,
        CancellationReason=types.SimpleNamespace(Error=1),
    ))

    ok = generate_audio.generate_audio_file("Hello world", str(tmp_file))
    assert ok is True
    assert DummySpeechConfig.last.authorization_token == 'fake-token-123'

    os.environ.pop('AZURE_ACCESS_TOKEN', None)
