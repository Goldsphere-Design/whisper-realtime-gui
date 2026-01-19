import argparse
import json
from datetime import timedelta
from pathlib import Path

import whisper


def format_timestamp(seconds: float) -> str:
    """Convert seconds to HH:MM:SS format."""
    td = timedelta(seconds=seconds)
    hours = td.seconds // 3600
    minutes = (td.seconds // 60) % 60
    secs = td.seconds % 60
    return f"{hours:02d}:{minutes:02d}:{secs:02d}"


def transcribe_file(model, audio_path: Path, input_root: Path, output_root: Path) -> None:
    result = model.transcribe(
        str(audio_path),
        language=None,  # auto-detect language; keeps whatever the audio is spoken in
        word_timestamps=True,
        condition_on_previous_text=False,
        no_speech_threshold=0.5,
        logprob_threshold=-1.0,
    )

    formatted_segments = []
    for segment in result["segments"]:
        formatted_segments.append(
            {
                "start_time": format_timestamp(segment["start"]),
                "end_time": format_timestamp(segment["end"]),
                "text": segment["text"].strip(),
            }
        )

    relative_path = audio_path.with_suffix("").relative_to(input_root)
    destination_dir = output_root / relative_path.parent
    destination_dir.mkdir(parents=True, exist_ok=True)

    base_name = destination_dir / relative_path.name

    with open(base_name.with_suffix(".json"), "w", encoding="utf-8") as f:
        json.dump(formatted_segments, f, ensure_ascii=False, indent=2)

    with open(base_name.with_suffix(".txt"), "w", encoding="utf-8") as f:
        for segment in formatted_segments:
            f.write(f'[{segment["start_time"]} -> {segment["end_time"]}]\n')
            f.write(f'{segment["text"]}\n\n')

    with open(base_name.with_suffix(".md"), "w", encoding="utf-8") as f:
        f.write("| Start | End | Text | Translation |\n")
        f.write("|-------|-----|------|-------------|\n")
        for segment in formatted_segments:
            f.write(f'| {segment["start_time"]} | {segment["end_time"]} | {segment["text"]} | |\n')

    try:
        from docx import Document

        doc = Document()
        table = doc.add_table(rows=1, cols=4)
        table.style = "Table Grid"

        header_cells = table.rows[0].cells
        header_cells[0].text = "Start"
        header_cells[1].text = "End"
        header_cells[2].text = "Text"
        header_cells[3].text = "Translation"

        for segment in formatted_segments:
            row_cells = table.add_row().cells
            row_cells[0].text = segment["start_time"]
            row_cells[1].text = segment["end_time"]
            row_cells[2].text = segment["text"]
            row_cells[3].text = ""

        doc.save(base_name.with_suffix(".docx"))
    except ImportError:
        print("python-docx not installed; skipping .docx export.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Transcribe all .mp3 files in a folder (recursively) using Whisper."
    )
    parser.add_argument(
        "input_root",
        type=Path,
        help="Folder containing .mp3 files (processed recursively).",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("transcripts"),
        help="Where to write transcripts (mirrors input folder structure).",
    )
    parser.add_argument(
        "--model",
        default="medium",
        help="Whisper model size/name to load (e.g. tiny, base, small, medium, large).",
    )
    args = parser.parse_args()

    input_root = args.input_root.expanduser().resolve()
    output_root = args.output.expanduser().resolve()

    mp3_files = sorted(p for p in input_root.rglob("*.mp3") if p.is_file())
    if not mp3_files:
        raise SystemExit(f"No .mp3 files found under {input_root}")

    print(f"Loading Whisper model '{args.model}' ...")
    whisper_model = whisper.load_model(args.model)

    print(f"Found {len(mp3_files)} .mp3 files. Writing outputs under {output_root}")
    for idx, audio_path in enumerate(mp3_files, start=1):
        print(f"[{idx}/{len(mp3_files)}] {audio_path}")
        transcribe_file(whisper_model, audio_path, input_root, output_root)

    print("Done.")
