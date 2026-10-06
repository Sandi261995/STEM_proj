# Prompt Strategies for Allergy-Aware Grocery Recommendations

Research materials for the ENGE817 Research Methods project *"How do different prompt-engineering strategies affect the safety and relevance of LLM-generated grocery product recommendations for users with food allergies?"* (Nway Sandi Oo, supervised by Dr. Weihua Li, October 2026).

The repository contains the product dataset, the four prompt templates, the experiment and scoring code, and the raw and scored outputs of every pilot and the full experiment, so the figures in the report can be traced back to the exact prompts and model responses.

> **Not medical advice.** Outputs are research artifacts produced from a fixed product snapshot. People with food allergies should always check current packaging and seek qualified advice.

## Repository contents

| Path | Contents |
|---|---|
| `data/product_dataset_40_prompt_fields_paknsave.xlsx` | 40 products from PAK'nSAVE Online (8 in each of 5 categories), one column per field supplied to the model |
| `Prompt_Templates_and_Response_Format.md` | The four prompt templates and the required JSON response format, read by the code at run time |
| `Prompt Templates for the Grocery Recommendation Experiment.docx` | The same templates as a Word document, for reading |
| `experiment/run-experiment.ts` | Builds the prompts, calls the OpenAI API and records every output |
| `experiment/score-results.ts` | Scores a results folder against the five metrics and the supplementary exclusion-list check |
| `experiment/package.json`, `package-lock.json` | Node dependencies and npm scripts |
| `experiment/.env.example` | Template for the API key and optional model settings |
| `experiment/results/` | One folder per run (see [Results folders](#results-folders)) |

## Experiment design

- **Independent variable:** prompt strategy: `basic-zero-shot`, `safety-focused`, `few-shot`, `self-verification`.
- **Scenarios:** 10 synthetic requests (SC01–SC10) combining milk, egg, peanut, tree-nut, soy, gluten or sesame restrictions with price, sugar, preparation-time or dietary constraints. The scenario wording is defined in `run-experiment.ts`.
- **Input per prompt:** the scenario plus the 8 products in the matching category.
- **Repetition:** 3 runs per scenario–strategy cell, giving 10 × 4 × 3 = **120 outputs**.
- **Dependent variables:** allergen-related error rate, relevance (1–5), constraint satisfaction, hallucination rate and consistency. The accuracy of the `excludedProducts` list is reported as a supplementary measure that was defined after the pilot.

### Model settings (full experiment)

| Setting | Value |
|---|---|
| Model | `gpt-5.6-luna` (OpenAI Responses API) |
| Reasoning effort | `medium` |
| Max output tokens | `8000` |
| Temperature | Not sent. Sampling was left to the model/API default, and the effective value was not recorded |
| Run date | 21 September 2026 |

The proposal specified GPT-5.5. After the pilot, the model was changed to GPT-5.6-Luna because only that model showed any recommendation variability, which kept the consistency metric informative. The report discusses the selection bias this introduces.

## Results folders

| Folder | Purpose | Model | Scenarios | Outputs |
|---|---|---|---|---|
| `dry-run-2026-09-13T10-00-26-856Z` | Prompts rendered to disk; no API calls | – | SC01–SC10 | – |
| `pilot-2026-09-13T10-28-59-905Z` | Pilot 1 | `gpt-5.6-luna` | SC01, SC08 | 24 |
| `pilot-2026-09-14T03-28-47-841Z` | Model check | `gpt-5.5-2026-04-23` | SC01, SC08 | 24 |
| `pilot-2026-09-14T08-58-48-169Z` | Pilot 2 (harder scenarios) | `gpt-5.6-luna` | SC04, SC10 | 24 |
| `experiment-2026-09-21T08-40-11-515Z` | Full experiment | `gpt-5.6-luna` | SC01–SC10 | 120 |

Each run folder contains:

- `manifest.json`: model settings, scenarios, the exact product snapshot sent to the model, and SHA-256 checksums of the dataset and template files
- `runs.jsonl`: one record per API call, with the exact prompts, raw response, parsed JSON, token usage and latency
- `scoring.xlsx`: summaries, per-output scores, the ground truth and review lists (written by `score-results.ts`)
- `manual-scores.xlsx`: the manual relevance and hallucination judgements
- Spreadsheet and text summaries prepared for the report

## Summary of findings

- All four strategies had a 0% allergen-related error rate, 100% constraint satisfaction, a mean relevance of 5.0/5 and a 0% hallucination rate. No differences were observed on these four metrics.
- Mean consistency was 93.3% for basic zero-shot, 96.7% for safety-focused, 93.3% for few-shot and 100% for self-verification. All variation came from SC09 and SC10 and involved only alternative suitable products.
- In the supplementary check, 18 of 120 outputs listed at least one suitable product in `excludedProducts`: 7 basic zero-shot, 0 safety-focused, 3 few-shot and 8 self-verification. This difference is descriptive and was not statistically tested.

Relevance and hallucination were scored by a single researcher, with no second rater. See the report for the full analysis, the scoring rules and the limitations.

## Reproducing the experiment

Requirements: Node.js 22 or later and an OpenAI API key.

```bash
cd experiment
npm install
cp .env.example .env        # then set OPENAI_API_KEY in .env (never commit .env)
```

```bash
npm run dry-run                                   # render all prompts without calling the API
npm run pilot -- --scenarios SC01,SC08            # pilot 1 (24 calls)
npm run pilot                                     # pilot 2: SC04 and SC10 by default (24 calls)
npm run experiment                                # full experiment (120 calls)
npm run score -- results/<results folder>         # score any results folder
```

To repeat the GPT-5.5 model check, set `OPENAI_MODEL=gpt-5.5-2026-04-23` in `.env` before running pilot 1.

Useful options for `run-experiment.ts`:

| Option | Meaning |
|---|---|
| `--scenarios SC01,SC08` | Run a subset of scenarios |
| `--strategies few-shot,safety-focused` | Run a subset of prompt strategies |
| `--runs 3` | Repeated runs per scenario–strategy cell |
| `--concurrency 1` | Parallel API calls |
| `--out results/<folder>` | Resume an interrupted run in an existing folder |

Optional `.env` overrides are `OPENAI_MODEL`, `REASONING_EFFORT`, `MAX_OUTPUT_TOKENS` and `TEMPERATURE`. Leave `TEMPERATURE` unset to match the reported runs.

The full run used about 316,000 input tokens and 113,000 output tokens, took about 17 minutes and cost about US$0.20. Model outputs are non-deterministic, and the `gpt-5.6-luna` alias does not expose a dated version, so a re-run may not reproduce the recorded outputs exactly. The recorded `runs.jsonl` files are the reference data for the report.

### Scoring

`score-results.ts` checks every output against the product snapshot stored in that run's `manifest.json`, not against the current workbook. It writes `scoring.xlsx` on every run. It creates `manual-scores.xlsx` once, for human judgement of relevance and hallucination, and then reads it back without ever overwriting it. Scoring rules that the proposal did not fully specify (for example, treating "Not declared" as safe, and "Under NZ$X" meaning strictly less than X) are listed in Section 4.6 of the report.

## Data source

Product information (names, prices, ingredients, allergen statements and nutrition or dietary labels) was collected manually from PAK'nSAVE Online in August 2026. It is a fixed research snapshot, and current products and labels may differ. No personal or clinical data was used; all scenarios are synthetic.
