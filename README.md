# Gemini Placement Intelligence Demos — ST. Peter's

Three demos show the progression from an LLM call to one autonomous agent and then a coordinated agent team.

## The demos

### 1. Resume–job description gap analysis

One Gemini call compares evidence in a resume with a target job. It returns a match score, demonstrated strengths, hiring gaps and honest resume improvements.

```bash
python demo_1_llm_call.py
```

### 2. Placement preparation agent

One Gemini agent analyses the resume, identifies priority gaps, creates a 14-day learning plan and calls a tool that generates an importable `.ics` study calendar.

```bash
python demo_2_single_agent.py --days 14
```

Generated calendar: `outputs/placement_preparation.ics`

### 3. Placement intelligence system

A sequential agent team collaborates:

```text
Job Researcher → Resume Analyst → Skills Coach → Interviewer → Career Manager
```

The Job Researcher uses Gemini Google Search grounding for current role expectations. Use `--no-live-search` when search grounding is unavailable.

```bash
python demo_3_multi_agent.py
python demo_3_multi_agent.py --no-live-search
```

## Setup

```bash
cd /Users/ritishjogi/Downloads/st-peters
source .venv/bin/activate
pip install -r requirements.txt
```

Paste your Gemini key into `.env`:

```dotenv
GEMINI_API_KEY=your_key_here
```

The key file is ignored by Git.

## Use your own resume and job description

```bash
python demo_1_llm_call.py --resume /path/resume.txt --job /path/job.txt
python demo_2_single_agent.py --resume /path/resume.txt --job /path/job.txt --days 21
python demo_3_multi_agent.py --resume /path/resume.txt --job /path/job.txt
```

Sample inputs are provided in `data/`.

## Low-token configuration

The default model is `gemini-3.5-flash-lite`. Each stage uses compact context, low temperature and a strict output cap:

```dotenv
LLM_MAX_OUTPUT_TOKENS=320
AGENT_MAX_OUTPUT_TOKENS=450
SPECIALIST_MAX_OUTPUT_TOKENS=240
MANAGER_MAX_OUTPUT_TOKENS=420
```

Every demo prints input, output and total token usage. The multi-agent demo also reports the approximate total across all five calls.
