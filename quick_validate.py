#!/usr/bin/env python3
"""Quick validation — check imports and API connectivity."""
import sys

def check_imports():
    try:
        from pptx import Presentation
        import requests
        print("✅ python-pptx, requests OK")
        return True
    except ImportError as e:
        print(f"❌ Missing dependency: {e}")
        print("   Run: pip install python-pptx requests")
        return False

def check_api(api_key: str):
    import requests
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    payload = {
        "model": "solar-mini",
        "messages": [{"role": "user", "content": "안녕"}],
        "max_tokens": 10,
    }
    try:
        r = requests.post(
            "https://api.upstage.ai/v1/solar/chat/completions",
            json=payload, headers=headers, timeout=15
        )
        r.raise_for_status()
        print(f"✅ Upstage API OK — model: solar-mini")
        return True
    except Exception as e:
        print(f"❌ API error: {e}")
        return False

if __name__ == "__main__":
    api_key = sys.argv[1] if len(sys.argv) > 1 else None
    ok = check_imports()
    if ok and api_key:
        check_api(api_key)
    elif not api_key:
        print("ℹ️  Pass API key to test: python quick_validate.py <api_key>")
