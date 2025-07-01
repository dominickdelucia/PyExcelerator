import os

try:
    import openai
except ImportError:
    openai = None


def excel_formula_to_python(formula: str) -> str:
    """Translate an Excel formula to Python code using an LLM if possible."""
    api_key = os.environ.get("OPENAI_API_KEY")
    if openai and api_key:
        openai.api_key = api_key
        prompt = (
            "Translate this Excel formula to an equivalent Python expression. "
            "Return only the Python expression.\n"
            f"Formula: {formula}\nPython:"
        )
        try:
            resp = openai.Completion.create(
                engine="text-davinci-003",
                prompt=prompt,
                max_tokens=60,
            )
            return resp.choices[0].text.strip()
        except Exception as e:
            return f"# LLM translation error: {e!r}"
    return f"# LLM translation unavailable. Formula: {formula}"
