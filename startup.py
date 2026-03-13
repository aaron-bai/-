"""
startup.py
Entry point for the automated paper review system.

Usage
-----
python startup.py \\
    --pdf path/to/paper.pdf \\
    --api-url https://api.openai.com/v1 \\
    --api-token sk-... \\
    --model gpt-4o \\
    --temperature 0.7 \\
    [--responder-model gpt-4o] \\
    [--responder-temperature 0.7] \\
    [--format zh|en|path/to/format.json] \
    [--output-path review_output.txt]

If --responder-model is not provided, the same model is used for all steps.
"""

import argparse
import os
import sys
from dataclasses import dataclass

from pdf_parser import PDFParser
from llm_client import LLMClient
from discipline_detector import DisciplineDetector
from docx_exporter import export_review_documents
from review_generator import ReviewGenerator
from review_responder import ReviewResponder


@dataclass
class ReviewPipelineResult:
    discipline: str
    review_questions: str
    review_responses: str
    output: str


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description="Automatically generate peer-review questions and author responses "
                    "for an academic paper using LLMs.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument(
        "--pdf",
        required=True,
        help="Path to the academic paper PDF file.",
    )
    parser.add_argument(
        "--api-url",
        required=True,
        help="Base URL of the OpenAI-compatible LLM API "
             "(e.g., https://api.openai.com/v1).",
    )
    parser.add_argument(
        "--api-token",
        required=True,
        help="API token / key for authenticating with the LLM service.",
    )
    parser.add_argument(
        "--model",
        required=True,
        help="Name of the primary LLM model to use for discipline detection and "
             "review question generation.",
    )
    parser.add_argument(
        "--temperature",
        type=float,
        default=0.7,
        help="Sampling temperature for the primary LLM (0.0 = deterministic, "
             "1.0 = creative).",
    )
    parser.add_argument(
        "--responder-model",
        default=None,
        help="Name of the LLM model used for generating review responses. "
             "Defaults to the value of --model.",
    )
    parser.add_argument(
        "--responder-temperature",
        type=float,
        default=None,
        help="Sampling temperature for the responder LLM. "
             "Defaults to the value of --temperature.",
    )
    parser.add_argument(
        "--language",
        default="English",
        help="Language in which to generate the review questions and responses.",
    )
    parser.add_argument(
        "--review-aspects",
        default="methodology",
        help="Specific aspects to focus on when generating the review questions "
             "(e.g., methodology, clarity, novelty). Provide as a newline-separated list.",
    )
    parser.add_argument(
        "--output-dir",
        required=True,
        help="Directory where the generated review and response Word documents will be saved.",
    )
    parser.add_argument(
        "--format",
        default="zh",
        help="DOCX format preset or custom JSON path: zh=中文学术风格, "
             "en=English academic style, or /path/to/format.json.",
    )
    parser.add_argument(
        "--output-path",
        default=None,
        help="Optional path to write the combined plain-text output.",
    )

    return parser.parse_args(argv)


def run(args) -> ReviewPipelineResult:
    """
    Execute the full review pipeline and return the combined output string.

    Parameters
    ----------
    args : argparse.Namespace
        Parsed CLI arguments (or any object with the same attributes).

    Returns
    -------
    ReviewPipelineResult
        Structured review data and combined textual output.
    """
    # ------------------------------------------------------------------ #
    # 1. Parse the PDF
    # ------------------------------------------------------------------ #
    print("[1/4] Parsing PDF…", flush=True)
    parser = PDFParser(args.pdf)
    paper_text = parser.extract_text()
    print(f"      Extracted {len(paper_text):,} characters from {args.pdf}")

    # ------------------------------------------------------------------ #
    # 2. Detect academic discipline
    # ------------------------------------------------------------------ #
    print("[2/4] Detecting academic discipline…", flush=True)
    primary_llm = LLMClient(
        api_url=args.api_url,
        api_token=args.api_token,
        model=args.model,
        temperature=args.temperature,
    )
    detector = DisciplineDetector(primary_llm)
    discipline = detector.detect(paper_text)
    print(f"      Detected discipline: {discipline}")

    # ------------------------------------------------------------------ #
    # 3. Generate review questions
    # ------------------------------------------------------------------ #
    print("[3/4] Generating review questions (problems only)…", flush=True)
    generator = ReviewGenerator(primary_llm)
    review_questions = generator.generate(paper_text, discipline, args.language, args.review_aspects)

    # ------------------------------------------------------------------ #
    # 4. Generate author responses to the review questions
    # ------------------------------------------------------------------ #
    print("[4/4] Generating responses to review questions…", flush=True)
    responder_model = args.responder_model or args.model
    responder_temperature = (
        args.responder_temperature
        if args.responder_temperature is not None
        else args.temperature
    )
    responder_llm = LLMClient(
        api_url=args.api_url,
        api_token=args.api_token,
        model=responder_model,
        temperature=responder_temperature,
    )
    responder = ReviewResponder(responder_llm)
    review_responses = responder.respond(
        review_questions=review_questions,
        paper_text=paper_text,
        discipline=discipline,
        language=args.language,
    )

    # ------------------------------------------------------------------ #
    # Assemble output
    # ------------------------------------------------------------------ #
    output = (
        "=" * 70 + "\n"
        f"AUTOMATED PEER REVIEW — {args.pdf}\n"
        + "=" * 70 + "\n\n"
        f"Detected Discipline: {discipline}\n\n"
        + "-" * 70 + "\n"
        "REVIEW QUESTIONS (Problems Only)\n"
        + "-" * 70 + "\n"
        + review_questions + "\n\n"
        + "-" * 70 + "\n"
        "RESPONSES TO REVIEW QUESTIONS\n"
        + "-" * 70 + "\n"
        + review_responses + "\n"
        + "=" * 70 + "\n"
    )

    return ReviewPipelineResult(
        discipline=discipline,
        review_questions=review_questions,
        review_responses=review_responses,
        output=output,
    )


def main(argv=None):
    args = parse_args(argv)
    base_dir = os.path.dirname(os.path.abspath(__file__))
    format_map = {
        "zh": os.path.join(base_dir, "format_zh_academic.json"),
        "en": os.path.join(base_dir, "format_en_academic.json"),
    }
    if args.format in format_map:
        selected_format_path = format_map[args.format]
    else:
        selected_format_path = os.path.abspath(os.path.expanduser(args.format))

    if not os.path.isfile(selected_format_path):
        print(
            f"ERROR: format file not found: {selected_format_path}. "
            "Use zh/en or provide a valid JSON file path via --format.",
            file=sys.stderr,
        )
        sys.exit(1)

    try:
        result = run(args)
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)

    print("\n" + result.output)

    try:
        output_dir = args.output_dir
        review_docx_path, response_docx_path = export_review_documents(
            output_dir=output_dir,
            paper_path=args.pdf,
            discipline=result.discipline,
            review_questions=result.review_questions,
            review_responses=result.review_responses,
            format_path=selected_format_path,
        )
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)

    print(f"评审意见 Word 已保存到: {review_docx_path}")
    print(f"回复 Word 已保存到: {response_docx_path}")

    if args.output_path:
        with open(args.output_path, "w", encoding="utf-8") as fh:
            fh.write(result.output)
        print(f"Output written to: {args.output_path}")


if __name__ == "__main__":
    main()
