"""Demo 2: one placement agent analyses gaps and schedules preparation."""

from __future__ import annotations

import argparse
from datetime import date, timedelta

from gemini_demo_utils import client, config, model_name, token_limit, usage_line
from placement_utils import (
    DEFAULT_JOB,
    DEFAULT_RESUME,
    create_preparation_schedule,
    placement_context,
)


def main() -> None:
    parser = argparse.ArgumentParser(description="Create and schedule a placement preparation plan.")
    parser.add_argument("--resume", default=str(DEFAULT_RESUME))
    parser.add_argument("--job", default=str(DEFAULT_JOB))
    parser.add_argument("--days", type=int, default=14)
    parser.add_argument("--start", default=(date.today() + timedelta(days=1)).isoformat())
    args = parser.parse_args()

    prompt = placement_context(args.resume, args.job) + f"""

Goal: prepare this candidate for the role in {args.days} days, beginning {args.start}.
Analyse the evidence, identify the highest-impact gaps, create 4–7 concrete learning tasks,
and call create_preparation_schedule exactly once. Then summarize the plan and calendar result.
"""
    response = client().models.generate_content(
        model=model_name(),
        contents=prompt,
        config=config(
            max_tokens=token_limit("AGENT_MAX_OUTPUT_TOKENS", 450),
            system=(
                "You are a placement preparation agent. Prioritize achievable gaps for the target role. "
                "Use the calendar tool to take action. Never claim skills that the resume does not show."
            ),
            tools=[create_preparation_schedule],
        ),
    )

    print(f"\nDEMO 2 — PLACEMENT PREPARATION AGENT ({model_name()})\n")
    print(response.text.strip())
    print("\nAction: an importable calendar was created in outputs/placement_preparation.ics")
    print(usage_line(response))


if __name__ == "__main__":
    main()

