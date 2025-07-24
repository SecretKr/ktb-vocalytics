import os
import io
import re
import torch
from dotenv import load_dotenv
from transformers import WhisperProcessor, WhisperForConditionalGeneration
from pydub import AudioSegment
import torchaudio
import time

# === Input Audio File ===
audio_file = "data/2 personal_loan.wav"

# === Start Timer ===
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
    device_asr = torch.device("mps")
elif torch.cuda.is_available():
    device_asr = torch.device("cuda")
else:
    device_asr = torch.device("cpu")

# === Load ASR Model ===
print("Loading biodatlab Whisper model...")
from transformers import logging
logging.set_verbosity_error()

# model_name = "biodatlab/distill-whisper-th-large-v3"
# model_name = "biodatlab/whisper-th-large-v3-combined"
model_name = "biodatlab/whisper-th-large-v3"
processor = WhisperProcessor.from_pretrained(model_name)
model = WhisperForConditionalGeneration.from_pretrained(model_name)
model.to(device_asr)

# === Load Entire Audio File ===
audio = AudioSegment.from_wav(audio_file)
chunk_length_ms = 30 * 1000  # 30 seconds per chunk
chunks = [audio[i:i + chunk_length_ms] for i in range(0, len(audio), chunk_length_ms)]

# === Transcribe in Chunks ===
transcription = ""
for idx, chunk in enumerate(chunks):
    buffer = io.BytesIO()
    chunk.export(buffer, format="wav")
    buffer.seek(0)

    try:
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

        with torch.no_grad():
            predicted_ids = model.generate(
                input_features,
                max_new_tokens=400,
                repetition_penalty=1.15,
                do_sample=False,
                early_stopping=True
            )
        text = processor.batch_decode(predicted_ids, skip_special_tokens=True)[0]
        cleaned = clean_thai_text(text)
        transcription += f"{cleaned} "

        print(f"Chunk {idx + 1}/{len(chunks)} done.")

    except Exception as e:
        print(f"Error in chunk {idx + 1}: {e}")
        transcription += "[Transcription Error] "

# === Save and Print Result ===
os.makedirs("transcript", exist_ok=True)
with open("transcript/transcript.txt", "w", encoding="utf-8") as f:
    f.write(transcription.strip())

print("\n=== Final Transcript ===")
print(transcription.strip())

# === End Timer and Print Total Time ===
end_time = time.time()
total_time = end_time - start_time
print(f"\nTotal execution time: {total_time:.2f} seconds")
