"""Demo 1: direct resume-to-job gap analysis with one Gemini call."""

from __future__ import annotations

import argparse

from gemini_demo_utils import generate, model_name, token_limit, usage_line
from placement_utils import DEFAULT_JOB, DEFAULT_RESUME, placement_context


def main() -> None:
    parser = argparse.ArgumentParser(description="Compare a resume with a job description.")
    parser.add_argument("--resume", default=str(DEFAULT_RESUME))
    parser.add_argument("--job", default=str(DEFAULT_JOB))
    args = parser.parse_args()

    prompt = placement_context(args.resume, args.job) + """

Return this compact assessment:
1. Match score from 0–100 with one-sentence justification
2. Top 4 demonstrated matches, each with resume evidence
3. Top 4 gaps, ordered by hiring impact
4. Three resume improvements that do not invent experience
"""
    response = generate(
        prompt,
        system=(
            "You are an evidence-based technical recruiter. Treat missing evidence as a gap. "
            "Never invent candidate experience. Be specific and concise."
        ),
        max_tokens=token_limit("LLM_MAX_OUTPUT_TOKENS", 320),
    )

    print(f"\nDEMO 1 — RESUME–JOB GAP ANALYSIS ({model_name()})\n")
    print(response.text.strip())
    print(f"\n{usage_line(response)}")


if __name__ == "__main__":
    main()

