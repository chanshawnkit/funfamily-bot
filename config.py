import os


def validate_anthropic_env() -> tuple[str, str, str]:
    """Return validated Anthropic settings or fail before serving traffic."""
    api_key = os.getenv("ANTHROPIC_API_KEY", "").strip()
    model = os.getenv("ANTHROPIC_MODEL", "").strip()
    base_url = os.getenv("ANTHROPIC_BASE_URL", "https://api.anthropic.com").strip()
    missing = [
        name
        for name, value in (
            ("ANTHROPIC_API_KEY", api_key),
            ("ANTHROPIC_MODEL", model),
        )
        if not value
    ]
    if missing:
        raise RuntimeError(
            "Missing required deployment environment variable(s): " + ", ".join(missing)
        )
    return api_key, model, base_url
