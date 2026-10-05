"""Demo 3: a sequential placement-intelligence agent team."""

from __future__ import annotations

import argparse

from google.genai import types

from gemini_demo_utils import generate, model_name, token_limit, usage_line
from placement_utils import DEFAULT_JOB, DEFAULT_RESUME, placement_context, read_text


def run_agent(role: str, instruction: str, context: str, *, live_search: bool = False):
    tools = [types.Tool(google_search=types.GoogleSearch())] if live_search else None
    response = generate(
        context,
        system=f"You are the {role}. {instruction} Be evidence-based and concise. No introduction.",
        max_tokens=token_limit("SPECIALIST_MAX_OUTPUT_TOKENS", 240),
        tools=tools,
    )
    print(f"\n[{role}]\n{response.text.strip()}\n{usage_line(response)}")
    return response


def total_tokens(responses: list) -> int:
    total = 0
    for response in responses:
        usage = getattr(response, "usage_metadata", None)
        total += (getattr(usage, "total_token_count", 0) or 0) if usage else 0
    return total


def main() -> None:
    parser = argparse.ArgumentParser(description="Five-agent placement intelligence pipeline.")
    parser.add_argument("--resume", default=str(DEFAULT_RESUME))
    parser.add_argument("--job", default=str(DEFAULT_JOB))
    parser.add_argument("--no-live-search", action="store_true", help="Skip Google Search grounding")
    args = parser.parse_args()

    resume = read_text(args.resume)
    job = read_text(args.job)
    base = placement_context(args.resume, args.job)
    print(f"\nDEMO 3 — PLACEMENT INTELLIGENCE SYSTEM ({model_name()})")

    researcher = run_agent(
        "Job Researcher",
        "Identify current expectations for this role in India. Separate job-description facts from market signals.",
        f"Target job description:\n{job}\n\nReturn 6 prioritized role expectations.",
        live_search=not args.no_live_search,
    )
    analyst = run_agent(
        "Resume Analyst",
        "Compare demonstrated evidence with role expectations. Return strengths, missing evidence, and risks.",
        f"{base}\n\nJOB RESEARCH:\n{researcher.text}",
    )
    coach = run_agent(
        "Skills Coach",
        "Turn gaps into a prioritized 14-day plan with measurable deliverables.",
        f"JOB DESCRIPTION:\n{job}\n\nRESUME ANALYSIS:\n{analyst.text}",
    )
    interviewer = run_agent(
        "Interviewer",
        "Create five targeted interview questions and what a strong answer must demonstrate.",
        f"RESUME:\n{resume}\n\nJOB DESCRIPTION:\n{job}\n\nCOACHING PLAN:\n{coach.text}",
    )
    manager_prompt = f"""
JOB RESEARCHER:\n{researcher.text}

RESUME ANALYST:\n{analyst.text}

SKILLS COACH:\n{coach.text}

INTERVIEWER:\n{interviewer.text}

Create one final placement strategy containing: readiness score, top 3 actions,
14-day milestones, interview focus, and a clear apply-now vs prepare-first recommendation.
"""
    manager = generate(
        manager_prompt,
        system="You are the Career Manager. Reconcile the specialist outputs into one decisive plan.",
        max_tokens=token_limit("MANAGER_MAX_OUTPUT_TOKENS", 420),
    )
    print(f"\n[Career Manager — Final Strategy]\n{manager.text.strip()}\n{usage_line(manager)}")
    responses = [researcher, analyst, coach, interviewer, manager]
    print(f"\nApproximate workflow tokens across 5 calls: {total_tokens(responses)}")


if __name__ == "__main__":
    main()

