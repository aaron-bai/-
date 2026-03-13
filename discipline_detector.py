"""
discipline_detector.py
Uses an LLM to detect the academic discipline of a research paper.
"""

from llm_client import LLMClient


_SYSTEM_PROMPT = (
    "You are an academic librarian with broad expertise across all scientific and "
    "humanistic disciplines. Given the full text (or an excerpt) of a research paper, "
    "identify the primary academic discipline or field it belongs to. "
    "Reply with a single, concise discipline name (e.g., 'Computer Science', "
    "'Molecular Biology', 'Economics', 'Materials Science', etc.). "
    "Do not include any explanation or additional text."
)

_USER_TEMPLATE = (
    "Please identify the primary academic discipline of the following research paper.\n\n"
    "--- PAPER TEXT (may be truncated) ---\n"
    "{paper_text}\n"
    "--- END OF PAPER TEXT ---\n\n"
    "Discipline:"
)

# How many characters of the paper to send to the LLM for discipline detection.
# Sending the full paper is expensive; the abstract + introduction are usually enough.
_MAX_CHARS = 8000
_FALLBACK_DISCIPLINE = "Unknown"


class DisciplineDetector:
    """Detect the academic discipline of a research paper using an LLM."""

    def __init__(self, llm_client: LLMClient):
        self.llm = llm_client

    def detect(self, paper_text: str) -> str:
        """
        Return the primary academic discipline of the paper.

        Parameters
        ----------
        paper_text : str
            Full or partial text of the research paper.

        Returns
        -------
        str
            The detected discipline (e.g., 'Computer Science').
        """
        cleaned_text = (paper_text or "").strip()
        if not cleaned_text:
            return _FALLBACK_DISCIPLINE

        excerpt = cleaned_text[:_MAX_CHARS]
        user_prompt = _USER_TEMPLATE.format(paper_text=excerpt)
        discipline = self.llm.chat(
            system_prompt=_SYSTEM_PROMPT,
            user_prompt=user_prompt,
            max_tokens=64,
        )

        normalized = self._normalize_discipline(discipline)
        if normalized:
            return normalized

        # Retry once with a shorter excerpt and stricter instruction when providers return blank content.
        retry_discipline = self.llm.chat(
            system_prompt=(
                "Return exactly discipline name of the following paper excerpt.  "
                "No explanation, no punctuation beyond the name."
            ),
            user_prompt=(
                "Identify the primary academic discipline for this paper excerpt:\n\n"
                f"{excerpt}\n\n"
                "Answer with discipline only."
            ),
        )
        normalized_retry = self._normalize_discipline(retry_discipline)
        return normalized_retry or _FALLBACK_DISCIPLINE

    @staticmethod
    def _normalize_discipline(value: str) -> str:
        text = (value or "").strip()
        if not text:
            return ""

        first_line = text.splitlines()[0].strip()
        if ":" in first_line:
            prefix, suffix = first_line.split(":", 1)
            if prefix.strip().lower() in {"discipline", "field", "学科", "领域"}:
                first_line = suffix.strip()

        return first_line.strip(" \t\r\n\"'`。；;,.，")
