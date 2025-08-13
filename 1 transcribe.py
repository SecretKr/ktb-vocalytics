import io
import os
import re
import time

import pandas as pd
import torch
import torchaudio
from dotenv import load_dotenv
from pyannote.audio import Pipeline
from pydub import AudioSegment
from transformers import WhisperForConditionalGeneration, WhisperProcessor

# === Input Audio File ===
audio_file = "\data\2 personal_loan.wav\2 personal_loan.wav"

start_time = time.time()
if not os.path.exists(audio_file):
    raise FileNotFoundError(f"The audio file was not found at: {audio_file}")

# === Clean Thai Text ===
def clean_thai_text(text):
    if text == "[Transcription Error]":
        return text
    cleaned_text = re.sub(r'(?<=[\u0E00-\u0E7F])\s+(?=[\u0E00-\u0E7F])', '', text)
    cleaned_text = re.sub(r'\s+', ' ', cleaned_text).strip()
    return cleaned_text

# === Load Environment and HF Token ===
load_dotenv()
hf_token = os.getenv("HF_TOKEN")
if hf_token is None:
    raise ValueError("Hugging Face token not found. Please set the HF_TOKEN environment variable.")

# === Device Configuration ===
if torch.backends.mps.is_available():
    device_pyannote = "mps"
    device_asr = torch.device("mps")
elif torch.cuda.is_available():
    device_pyannote = "cuda"
    device_asr = torch.device("cuda")
else:
    device_pyannote = "cpu"
    device_asr = torch.device("cpu")

# === Diarization Pipeline ===
print("Starting speaker diarization...")
diarization_pipeline = Pipeline.from_pretrained(
    "pyannote/speaker-diarization-3.1",
    use_auth_token=hf_token
)
diarization_pipeline.to(torch.device(device_pyannote))
diarization = diarization_pipeline(audio_file)

# === Diarization DataFrame ===
data = [{
    'start': segment.start,
    'end': segment.end,
    'speaker': speaker
} for segment, _, speaker in diarization.itertracks(yield_label=True)]
diarization_df = pd.DataFrame(data)

# === Load ASR Model ===
print("Loading biodatlab Whisper model...")
from transformers import logging

logging.set_verbosity_error()  # Reduce warnings

# model_name = "biodatlab/distill-whisper-th-large-v3"
# model_name = "biodatlab/whisper-th-large-v3-combined"
model_name = "biodatlab/whisper-th-large-v3"
processor = WhisperProcessor.from_pretrained(model_name)
model = WhisperForConditionalGeneration.from_pretrained(model_name)
model.to(device_asr)

# === Load Audio for Segmentation ===
full_audio = AudioSegment.from_wav(audio_file)

transcribed_segments = []
for i, row in diarization_df.iterrows():
    print(f"🔄 Processing segment {i+1}/{len(diarization_df)}: {row['start']:.2f}s - {row['end']:.2f}s")
    start_time_ms = int(row['start'] * 1000)
    end_time_ms = int(row['end'] * 1000)
    segment_audio = full_audio[start_time_ms:end_time_ms]

    # Save to in-memory buffer
    buffer = io.BytesIO()
    segment_audio.export(buffer, format="wav")
    buffer.seek(0)

    try:
        # Load and preprocess audio
        waveform, sample_rate = torchaudio.load(buffer)

        if sample_rate != 16000:
            resampler = torchaudio.transforms.Resample(orig_freq=sample_rate, new_freq=16000)
            waveform = resampler(waveform)

        if waveform.shape[0] > 1:
            waveform = waveform.mean(dim=0, keepdim=True)

        input_features = processor(
            waveform.squeeze().numpy(),
            sampling_rate=16000,
            return_tensors="pt"
        ).input_features.to(device_asr)

        # Generate transcription
        with torch.no_grad():
            predicted_ids = model.generate(input_features)
        transcribed_text = processor.batch_decode(predicted_ids, skip_special_tokens=True)[0]
        cleaned_text = clean_thai_text(transcribed_text)

    except Exception as e:
        print(f"Error in segment {i}: {e}")
        cleaned_text = "[Transcription Error]"

    transcribed_segments.append({
        'start': row['start'],
        'end': row['end'],
        'speaker': row['speaker'],
        'text': cleaned_text
    })

# === Save and Print Results ===
final_transcript_df = pd.DataFrame(transcribed_segments)
final_transcript_df.to_csv("transcript/transcript.csv", index=False, encoding='utf-8')

print("\n=== Final Transcript ===")
for i, row in final_transcript_df.iterrows():
    print(f"[{row['start']:.2f}s - {row['end']:.2f}s] {row['speaker']}: {row['text']}")

# === End Timer and Print Total Time ===
end_time = time.time()
total_time = end_time - start_time
print(f"\nTotal execution time: {total_time:.2f} seconds")
