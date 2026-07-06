"""
Vision-AI provider switch.

AI_PROVIDER=ollama (default) - matches the previous hardcoded behavior:
    an OpenAI-compatible client pointed at OLLAMA_BASE_URL (local dev / Ollama Cloud).
AI_PROVIDER=gemini - uses Google's Gemini API (GEMINI_API_KEY), for
    deployments (e.g. the Oracle VM) where Ollama isn't reachable.

Both branches return a single plain-text string, so callers (e.g.
bom_generator.py's extract_isometric_view) don't need to know which
provider answered.
"""
import os

AI_PROVIDER = os.getenv("AI_PROVIDER", "ollama").strip().lower()

OLLAMA_BASE_URL = os.getenv("OLLAMA_BASE_URL", "http://localhost:11434/v1")
OLLAMA_API_KEY = os.getenv("OLLAMA_API_KEY", "ollama")
OLLAMA_VISION_MODEL = os.getenv("OLLAMA_VISION_MODEL", "gemma4:31b-cloud")

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")
GEMINI_VISION_MODEL = os.getenv("GEMINI_VISION_MODEL", "gemini-2.0-flash")


def _ollama_vision_completion(prompt: str, base64_image: str) -> str:
    from openai import OpenAI

    client = OpenAI(base_url=OLLAMA_BASE_URL, api_key=OLLAMA_API_KEY)
    response = client.chat.completions.create(
        model=OLLAMA_VISION_MODEL,
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{base64_image}"}},
                ],
            }
        ],
        temperature=0.0,
        top_p=0.1,
    )
    return response.choices[0].message.content


def _gemini_vision_completion(prompt: str, base64_image: str) -> str:
    import base64
    import google.generativeai as genai

    if not GEMINI_API_KEY:
        raise RuntimeError("AI_PROVIDER=gemini but GEMINI_API_KEY is not set.")

    genai.configure(api_key=GEMINI_API_KEY)
    model = genai.GenerativeModel(GEMINI_VISION_MODEL)
    response = model.generate_content(
        [
            prompt,
            {"mime_type": "image/png", "data": base64.b64decode(base64_image)},
        ],
        generation_config={"temperature": 0.0, "top_p": 0.1},
    )
    return response.text


def get_vision_completion(prompt: str, base64_image: str) -> str:
    """Single entry point used by callers regardless of the active provider."""
    print(f"   [VISION] Calling {AI_PROVIDER} for isometric extraction...", flush=True)
    if AI_PROVIDER == "gemini":
        return _gemini_vision_completion(prompt, base64_image)
    return _ollama_vision_completion(prompt, base64_image)
