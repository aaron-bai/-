"""
review_responder.py
Uses an LLM to draft author responses to each peer-review question.

For every question raised by the reviewer, the responder:
  1. Explains why the question / problem is valid or what misunderstanding it reflects.
  2. Proposes a concrete way to address or resolve the issue in a revised manuscript.
"""

import json
import re

from llm_client import LLMClient


_SYSTEM_PROMPT = (
    "You are an expert academic author responding to peer-review comments for a paper "
    "in the field of {discipline}. "
    "You have received a JSON array of reviewer issues. "
    "For EACH item, provide a concise summary of the problem and a concrete response.\n\n"
    "Return ONLY a JSON array. Do not include markdown, prose, or code fences. "
    "Each array item must be an object with exactly these keys:\n"
    "  - \"problem\": summary of the reviewer's issue\n"
    "  - \"responde\": response and concrete resolution plan\n\n"
    "Be professional, constructive, and specific."
    "You should write your responses in {language}."
)

_USER_TEMPLATE = (
    "Below is the reviewer JSON raised for a paper in the field of {discipline}.\n\n"
    "--- REVIEW JSON ---\n"
    "{review_questions}\n"
    "--- END OF REVIEW JSON ---\n\n"
    "For context, here is a summary of the paper:\n\n"
    "--- PAPER TEXT (may be truncated) ---\n"
    "{paper_text}\n"
    "--- END OF PAPER TEXT ---\n\n"
    "Return ONLY JSON. No extra text."
)

_MAX_PAPER_CHARS = 12000
_REQUIRED_KEYS = ("problem", "responde")
_CODE_FENCE_PATTERN = re.compile(r"^```(?:json)?\s*|\s*```$", re.IGNORECASE)


def _extract_json_payload(raw_text: str) -> str:
    text = raw_text.strip()
    text = _CODE_FENCE_PATTERN.sub("", text).strip()

    try:
        json.loads(text)
        return text
    except json.JSONDecodeError:
        pass

    start = text.find("[")
    end = text.rfind("]")
    if start != -1 and end != -1 and end > start:
        candidate = text[start:end + 1].strip()
        json.loads(candidate)
        return candidate

    raise json.JSONDecodeError("No JSON array found", text, 0)


def _normalize_response_items(raw_text: str) -> str:
    try:
        payload = _extract_json_payload(raw_text)
        data = json.loads(payload)
    except json.JSONDecodeError as exc:
        raise RuntimeError("LLM did not return valid responder JSON.") from exc

    if not isinstance(data, list):
        raise RuntimeError("Responder output must be a JSON array.")

    normalized: list[dict[str, str]] = []
    for index, item in enumerate(data, start=1):
        if not isinstance(item, dict):
            raise RuntimeError(f"Responder item #{index} is not a JSON object.")

        normalized_item: dict[str, str] = {}
        for key in _REQUIRED_KEYS:
            value = item.get(key, "")
            normalized_item[key] = str(value).strip()
        normalized.append(normalized_item)

    return json.dumps(normalized, ensure_ascii=False, indent=2)


class ReviewResponder:
    """Draft author responses to peer-review questions using an LLM."""

    def __init__(self, llm_client: LLMClient):
        self.llm = llm_client

    def respond(
        self,
        review_questions: str,
        paper_text: str,
        discipline: str,
        language: str = "English"
    ) -> str:
        """
        Return a numbered list of responses to the review questions.

        Parameters
        ----------
        review_questions : str
            The numbered list of review questions produced by ReviewGenerator.
        paper_text : str
            Full or partial text of the research paper (for context).
        discipline : str
            The academic discipline of the paper.
        language : str
            The language in which to generate the responses.


        Returns
        -------
        str
            Numbered response list with explanation and resolution for each question.
        """
        excerpt = paper_text[:_MAX_PAPER_CHARS]
        system_prompt = _SYSTEM_PROMPT.format(discipline=discipline, language=language)
        user_prompt = _USER_TEMPLATE.format(
            discipline=discipline,
            review_questions=review_questions,
            paper_text=excerpt,
            language=language
        )
        raw = self.llm.chat(
            system_prompt=system_prompt,
            user_prompt=user_prompt,
        ).strip()
        return _normalize_response_items(raw)
