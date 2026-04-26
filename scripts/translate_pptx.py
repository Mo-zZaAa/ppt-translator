#!/usr/bin/env python3
"""
PPT Translator — Korean → English
Uses Upstage Solar mini API for batch translation.
Preserves all formatting (font, size, color, position).
"""

import json
import re
import sys
import argparse
import requests
from pathlib import Path
from pptx import Presentation
from pptx.util import Pt

UPSTAGE_API_URL = "https://api.upstage.ai/v1/solar/chat/completions"
KOREAN_RE = re.compile(r"[가-힣ㄱ-ㅎㅏ-ㅣ]")


def has_korean(text: str) -> bool:
    return bool(KOREAN_RE.search(text))


def collect_runs(prs):
    """Collect all (run, original_text) tuples that contain Korean."""
    runs = []

    def _from_tf(tf):
        for para in tf.paragraphs:
            for run in para.runs:
                if run.text and has_korean(run.text):
                    runs.append(run)

    for slide in prs.slides:
        for shape in slide.shapes:
            _process_shape(shape, _from_tf)

    return runs


def _process_shape(shape, callback):
    from pptx.util import Emu
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    if shape.has_text_frame:
        callback(shape.text_frame)

    if shape.shape_type == MSO_SHAPE_TYPE.TABLE:
        for row in shape.table.rows:
            for cell in row.cells:
                callback(cell.text_frame)

    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        for child in shape.shapes:
            _process_shape(child, callback)


def _translate_chunk(chunk: list[str], api_key: str) -> dict[str, str]:
    """Translate a single chunk via one API call."""
    system_prompt = (
        "You are a professional Korean-to-English translator specializing in business and consulting documents. "
        "The user will send a JSON array of Korean strings. "
        "Return a JSON object where each key is the EXACT original Korean string and the value is the English translation. "
        "Rules: keep proper nouns, company names, numbers, and English words unchanged. "
        "Return ONLY the JSON object, no markdown, no extra text."
    )
    payload = {
        "model": "solar-mini",
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": json.dumps(chunk, ensure_ascii=False)},
        ],
        "response_format": {"type": "json_object"},
        "temperature": 0.1,
        "max_tokens": 4096,
    }
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    resp = requests.post(UPSTAGE_API_URL, json=payload, headers=headers, timeout=60)
    resp.raise_for_status()
    content = resp.json()["choices"][0]["message"]["content"]
    return json.loads(content)


def batch_translate(texts: list[str], api_key: str, chunk_size: int = 50) -> dict[str, str]:
    """Translate all Korean strings, chunked to avoid timeouts."""
    unique = list(dict.fromkeys(texts))  # deduplicate, preserve order
    result = {}

    chunks = [unique[i:i + chunk_size] for i in range(0, len(unique), chunk_size)]
    print(f"Splitting {len(unique)} unique strings into {len(chunks)} chunks of ≤{chunk_size}...")

    for i, chunk in enumerate(chunks, 1):
        print(f"  Translating chunk {i}/{len(chunks)} ({len(chunk)} strings)...", end=" ", flush=True)
        translated = _translate_chunk(chunk, api_key)
        result.update(translated)
        print("done")

    return result


def translate_pptx(input_path: str, output_path: str, api_key: str) -> int:
    """Main translation function. Returns number of replaced runs."""
    prs = Presentation(input_path)

    runs = collect_runs(prs)
    if not runs:
        print("No Korean text found.")
        prs.save(output_path)
        return 0

    texts = [r.text for r in runs]
    print(f"Found {len(texts)} Korean runs ({len(set(texts))} unique). Calling API...")

    translation_map = batch_translate(texts, api_key)

    replaced = 0
    for run in runs:
        original = run.text
        translated = translation_map.get(original)
        if translated and translated != original:
            run.text = translated
            replaced += 1

    prs.save(output_path)
    print(f"Translated {replaced}/{len(runs)} runs. Saved → {output_path}")
    return replaced


def main():
    parser = argparse.ArgumentParser(description="Translate Korean PPT to English")
    parser.add_argument("input", help="Input .pptx file path")
    parser.add_argument("-o", "--output", help="Output file path (default: input_EN.pptx)")
    parser.add_argument("--api-key", required=True, help="Upstage API key")
    args = parser.parse_args()

    input_path = Path(args.input)
    output_path = args.output or str(input_path.with_stem(input_path.stem + "_EN"))

    translate_pptx(str(input_path), output_path, args.api_key)


if __name__ == "__main__":
    main()
