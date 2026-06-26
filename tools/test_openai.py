from __future__ import annotations

import os
import sys
from pathlib import Path

from dotenv import load_dotenv
from openai import APIConnectionError, APITimeoutError, OpenAI, OpenAIError


PROJECT_DIR = Path(__file__).resolve().parents[1]
DEFAULT_MODEL = "gpt-4.1-mini"


def main() -> int:
    load_dotenv(PROJECT_DIR / ".env")

    api_key = os.getenv("OPENAI_API_KEY", "").strip()
    if not api_key:
        print("ERROR: OPENAI_API_KEY no está configurada en .env", file=sys.stderr)
        return 1

    client = OpenAI(api_key=api_key, timeout=30)
    model = os.getenv("OPENAI_MODEL", DEFAULT_MODEL).strip() or DEFAULT_MODEL

    try:
        response = client.responses.create(
            model=model,
            input="Responde únicamente OK",
        )
    except (APIConnectionError, APITimeoutError) as exc:
        print(f"ERROR: no se pudo conectar con OpenAI: {exc}", file=sys.stderr)
        return 2
    except OpenAIError as exc:
        print(f"ERROR: OpenAI rechazó la solicitud: {exc}", file=sys.stderr)
        return 3

    print(response.output_text.strip())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
