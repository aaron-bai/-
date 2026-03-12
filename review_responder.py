"""
review_responder.py
Uses an LLM to draft author responses to each peer-review question.

For every question raised by the reviewer, the responder:
  1. Explains why the question / problem is valid or what misunderstanding it reflects.
  2. Proposes a concrete way to address or resolve the issue in a revised manuscript.
"""

from llm_client import LLMClient


_SYSTEM_PROMPT = (
    "You are an expert academic author responding to peer-review comments for a paper "
    "in the field of {discipline}. "
    "You have received a numbered list of reviewer questions and concerns. "
    "For EACH question, you must:\n"
    "  1. Briefly explain the nature of the problem or what the reviewer's concern is "
    "pointing to.\n"
    "  2. Propose a specific, actionable way to address or resolve the issue (e.g., "
    "additional experiments, clarified methodology, added discussion, corrected figures, "
    "etc.).\n\n"
    "Format your response as a numbered list that mirrors the reviewer's list. "
    "Each item should follow this structure:\n"
    "  [Question N] Explanation: <explain the problem>\n"
    "               Resolution: <how to address it>\n\n"
    "Be professional, constructive, and specific."
    "You should write your responses in {language}."
)

_USER_TEMPLATE = (
    "Below are the peer-review questions raised for a paper in the field of {discipline}.\n\n"
    "--- REVIEW QUESTIONS ---\n"
    "{review_questions}\n"
    "--- END OF REVIEW QUESTIONS ---\n\n"
    "For context, here is a summary of the paper:\n\n"
    "--- PAPER TEXT (may be truncated) ---\n"
    "{paper_text}\n"
    "--- END OF PAPER TEXT ---\n\n"
    "Please provide a response to each review question:"
)

_MAX_PAPER_CHARS = 12000


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
        return self.llm.chat(
            system_prompt=system_prompt,
            user_prompt=user_prompt,
        ).strip()
