# PPT Translator Skill

Translate Korean PowerPoint files to English using Upstage Solar API.
Preserves all formatting — fonts, colors, sizes, positions, images untouched.

## Trigger Phrases
- "이 PPT 번역해줘"
- "PPT 영어로 바꿔줘"
- "한→영 번역"
- "영문 버전 만들어줘"
- "translate this PPT"
- "PPT to English"

## When to Use
When the user provides a `.pptx` file and asks for Korean→English translation.

## How to Run

```bash
pip install python-pptx requests
python scripts/translate_pptx.py <input.pptx> --api-key <UPSTAGE_API_KEY>
```

Output file: `<input>_EN.pptx` in the same directory.

## What It Does
1. Loads the .pptx with python-pptx
2. Scans all shapes (text frames, tables, group shapes) for Korean text
3. Collects all Korean runs → deduplicates → sends to Upstage Solar mini in **one API call**
4. Replaces only `run.text` — all formatting preserved
5. Saves the translated file

## What It Does NOT Change
- Images, charts, diagrams
- Font family, size, color, bold/italic
- Shape position, size, layout
- Non-Korean text (English, numbers, proper nouns)

## Example
```bash
python scripts/translate_pptx.py ~/Desktop/보고서.pptx \
  --api-key up_t78SMLgcCwQFAVMrwCZ57Tfn51qSW \
  -o ~/Desktop/Report_EN.pptx
```
