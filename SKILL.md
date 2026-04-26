# PPT Translator Skill

Translate Korean PowerPoint files to English using Upstage Solar API.
All formatting is preserved — fonts, colors, sizes, positions, images untouched.

---

## Why This Skill Exists

### Without this skill:
| Problem | What happens |
|---------|-------------|
| **Claude token waste** | Asking Claude to translate a PPT directly burns large amounts of Claude tokens — a 30-slide deck can cost thousands of tokens |
| **Format destruction** | Claude extracts text, translates, but loses all table structures, text box positions, font styles |
| **Manual work** | Every translation requires: copy text → translate → manually re-paste into each shape |

### With this skill:
| Benefit | How |
|---------|-----|
| **Zero Claude tokens on translation** | All translation offloaded to Upstage Solar mini API |
| **100% format preservation** | `python-pptx` directly modifies `run.text` — positions, fonts, images, tables untouched |
| **One command automation** | Say "이 PPT 번역해줘" → done |

---

## Trigger Phrases
- "이 PPT 번역해줘"
- "PPT 영어로 바꿔줘"
- "한→영 번역"
- "영문 버전 만들어줘"
- "translate this PPT"
- "PPT to English"
- "영어로 번역해서 저장해줘"

---

## How to Run

```bash
pip install python-pptx requests
python scripts/translate_pptx.py <input.pptx> --api-key <UPSTAGE_API_KEY>
```

Output: `<input>_EN.pptx` in same directory (or specify with `-o`).

### Example
```bash
python scripts/translate_pptx.py ~/Desktop/보고서.pptx \
  --api-key up_xxxx \
  -o ~/Desktop/Report_EN.pptx
```

---

## How It Works

```
PPT 로드
  ↓
모든 shape 순회 (일반 shape + 테이블 셀 + 그룹 shape 재귀)
  ↓
paragraph 단위로 한국어 텍스트 수집
  (run들을 합쳐서 완전한 문장으로 복원)
  ↓
중복 제거 후 Upstage Solar mini API 배치 호출
  (40개 paragraph씩 청크 처리)
  ↓
번역 결과 적용:
  - 첫 번째 run에 번역 전체 삽입
  - 나머지 run은 빈 문자열로 (포맷 속성은 유지)
  ↓
저장
```

### Why paragraph-level (not run-level)?

PPT 내부에서 "앰버서더 2기 김세은" 한 문장이 실제로 이렇게 쪼개져 저장됩니다:

```
run[0] = "앰버서더 "   (나눔고딕 14pt)
run[1] = "2"           (Arial 14pt — 숫자라 폰트가 다름)
run[2] = "기 "         (나눔고딕 14pt)
run[3] = "김세은"      (나눔고딕 14pt bold)
```

run 단위로 번역하면 "기" 혼자 → "Issue", "2" → "2" 처럼 맥락 없이 깨집니다.
paragraph 단위로 합쳐서 번역하면 전체 문맥이 유지됩니다.

---

## What Is NOT Changed
- Images, charts, SmartArt, diagrams
- Shape positions (left, top, width, height)
- Font size, color, bold/italic (first run's style covers the paragraph)
- Table row/column structure
- Non-Korean text (English, numbers, proper nouns per translation rules)
- Slide backgrounds and layouts

---

## Requirements
- Python 3.8+
- `pip install python-pptx requests`
- Upstage API key from https://console.upstage.ai
