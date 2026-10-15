[README.md](https://github.com/user-attachments/files/33485868/README.md)
# ToR Review Assistant

**This README was generated with AI!**

**Turn reviewer feedback into a reusable checklist and evidence-linked reviews of evaluation Terms of Reference.**

This project contains two Python notebooks for supporting the review of evaluation Terms of Reference (ToRs). The first extracts reviewer comments from annotated Word documents and generates a draft checklist. The second uses a **human-reviewed checklist**, document parsing, and large language models (LLMs) to assess a new draft ToR and produce revision guidance.

The workflow is organised around the SECO ToR section structure used in the notebooks. It is a notebook-based research prototype, not an autonomous approval system or a packaged application.

> **Human review is essential.** Generated checklist items, classifications, and suggested wording require expert verification. The assistant creates a separate review report; it does not edit the original ToR or apply tracked changes.

## Contents

[Workflow](#workflow) · [Setup](#setup) · [Files and credentials](#files-and-credentials) · [Generate a checklist](#generate-a-checklist) · [Prepare the final checklist](#prepare-the-final-checklist) · [Review a draft ToR](#review-a-draft-tor) · [Outputs](#outputs) · [Models and request volume](#models-and-request-volume) · [Privacy](#privacy-and-responsible-use) · [Limitations and troubleshooting](#limitations-and-troubleshooting) · [Reproducibility](#reproducibility-and-quality-checks) · [Contributing](#contributing) · [License](#license-and-attribution)

## Workflow

| Notebook | Purpose | Main inputs | Main outputs |
| --- | --- | --- | --- |
| [`Checklist_Generation.ipynb`](Checklist_Generation.ipynb) | Extract reviewer feedback, clean the evidence, and generate reusable checklist items. | Commented `.docx` ToRs. | Cleaned comment dataset and draft checklist CSV. |
| [`ToR_Review_Assistant.ipynb`](ToR_Review_Assistant.ipynb) | Check a draft ToR against selected fail signals and prepare revision guidance. | Human-reviewed checklist CSV, draft `.docx` ToR, and report logo. | Excel results and a Word review report. |

```text
Commented evaluation ToRs (.docx)
    |
    v
Checklist_Generation.ipynb
    |-- ToR_comments_table_clean.csv
    `-- ToR_Checklist_Before_Human_Feedback.csv
                    |
                    v
          Human review and schema adaptation
                    |
                    v
             Final_Checklist.csv + draft ToR (.docx)
                    |
                    v
          ToR_Review_Assistant.ipynb
                    |-- ToR_Review_Data.xlsx
                    `-- ToR_Review_Report.docx
```

**Already have a compatible, reviewed checklist?** Start with `ToR_Review_Assistant.ipynb`; checklist generation does not need to be repeated for each draft ToR.

> **The notebook handoff is not automatic.** The generated checklist and the review assistant use different CSV schemas. Renaming the generated file to `Final_Checklist.csv` is not sufficient. Follow [Prepare the final checklist](#prepare-the-final-checklist).

<details>
<summary>Section structure</summary>

Both notebooks define the following section buckets, although their keyword rules differ:

| Section | Title |
| --- | --- |
| 1 | Purpose of the Evaluation |
| 2 | Description of the Intervention |
| 3 | Scope and Focus of the Evaluation / Evaluation Questions |
| 4 | Evaluation Methods and Process |
| 5 | Deliverables |
| 6 | Schedule |
| 7 | Evaluation Team / Qualifications |
| 8 | Budget |
| 9 | Annexes |

Unmatched evidence is initially assigned to `(unknown)`. During checklist generation, non-numbered bucket labels are normalised to `Overarching`. In the review parser, unmatched content remains `(unknown)`. Generation also normalises separators in section titles, so exported labels can differ slightly from the original bucket names.

</details>

## Setup

### Requirements

Use Google Colab, as configured in the notebooks, or a local Jupyter environment after adapting the Drive and filesystem cells. The saved installation logs show Python 3.12; no dependency lockfile or supported-version matrix is supplied.

You also need an OpenAI API key with access to the configured models, permission to process the input documents, and the private files described below. Model access, billing, and request limits must be checked for your own API account.

### Install dependencies

The following suggested command consolidates the libraries imported by both notebooks and adds `openpyxl` for Excel export. The notebooks' original installation cells do **not** explicitly install every dependency.

In a notebook cell:

```python
%pip install pandas spacy ftfy python-dotenv openai "pydantic>=2,<3" matplotlib python-docx openpyxl
!python -m spacy download xx_ent_wiki_sm
```

Restart the kernel/runtime if prompted, then run the imports again. Pydantic 2 is used for the checklist's structured-output schema. The spaCy model must be installed separately from the `spacy` package.

<details>
<summary>Local Jupyter setup</summary>

Run these commands from your project directory. The commands are setup examples, not a tested dependency lockfile.

```bash
python -m venv .venv

# macOS / Linux
source .venv/bin/activate

# Windows PowerShell: use this activation command instead
# .venv\Scripts\Activate.ps1

python -m pip install jupyterlab pandas spacy ftfy python-dotenv openai "pydantic>=2,<3" matplotlib python-docx openpyxl
python -m spacy download xx_ent_wiki_sm
python -m jupyterlab
```

Skip the `google.colab` import and `drive.mount(...)` cell in **both** notebooks. Replace the `/content/drive/...` paths with local paths. For example, use `PROJECT_ROOT = str(Path.cwd())` after importing `Path`, provided Jupyter's working directory is the project root.

</details>

## Files and credentials

The supplied project materials contain the two notebooks, not their referenced datasets, final checklist, draft ToRs, credentials, or logo. Provide your own authorised files. Saved cell outputs are not substitutes for those inputs and may themselves contain sensitive excerpts.

A suggested working directory is shown below. This is a layout to create, not a list of additional files bundled with the notebooks.

```text
your-project/
|-- README.md
|-- Checklist_Generation.ipynb
|-- ToR_Review_Assistant.ipynb
|-- .env                                      # Private; do not commit
|-- Checklist_Documents/
|   |-- commented_tors/                       # Your commented .docx inputs
|   |-- ToR_comments_table_clean.csv          # Generated
|   |-- ToR_Checklist_Before_Human_Feedback.csv # Generated
|   `-- Final_Checklist.csv                   # Human-reviewed and adapted
|-- Input/
|   |-- draft_tor.docx                        # Your draft ToR
|   `-- Logo.jpg                              # Required by unmodified report cell
`-- Output/
    |-- ToR_Review_Data.xlsx                   # Generated
    `-- ToR_Review_Report.docx                 # Generated
```

Create a private `.env` file at the configured project root:

```dotenv
OPENAI_API_KEY=replace_with_your_api_key
```

The notebooks load this file through `ENV_FILE_PATH` or `ENV_PATH` and read `OPENAI_API_KEY`. Model names are configured in code, not through this environment variable.

For Colab, mount Google Drive using the existing cell and place your working files there. Replace the original author-specific paths before running the processing cells. The configuration examples below use `/content/drive/MyDrive/ToR_Review_Project`.

## Generate a checklist

Open `Checklist_Generation.ipynb`. After setup and Drive mounting, replace its path-configuration cell with paths appropriate to your environment. For the suggested layout:

```python
from pathlib import Path

PROJECT_ROOT = "/content/drive/MyDrive/ToR_Review_Project"
CHECKLIST_OUT_DIR = f"{PROJECT_ROOT}/Checklist_Documents"
SECO_TOR_DIR = f"{CHECKLIST_OUT_DIR}/commented_tors"
CLEANED_COMMENTS_CSV_PATH = f"{CHECKLIST_OUT_DIR}/ToR_comments_table_clean.csv"
ENV_FILE_PATH = f"{PROJECT_ROOT}/.env"
CHECKLIST_OUT_PATH = f"{CHECKLIST_OUT_DIR}/ToR_Checklist_Before_Human_Feedback.csv"

Path(CHECKLIST_OUT_DIR).mkdir(parents=True, exist_ok=True)
```

Place your annotated `.docx` files in `SECO_TOR_DIR`; the input scan is not recursive. Documents must contain actual Word comments with anchors, not just tracked changes or inline reviewer notes. Files without `word/comments.xml` are skipped, and the pipeline stops when no anchored comments are extracted.

Run the cells in order, inspecting the extraction and cleaning results before the model-generation loop. The notebook links comments to anchor paragraphs and heading context, assigns section buckets using keywords, repairs text encoding and spacing, and applies partial anonymisation. It then groups and samples evidence, requests structured checklist items, removes duplicate requirements, and exports the results.

### Generation settings

| Setting | Default in the notebook | Meaning |
| --- | --- | --- |
| `MODEL` | `gpt-5.2` | Checklist-generation model identifier. |
| `TARGET_ITEMS_PER_SECTION` | `10` | Requested number of draft items for each represented section. |
| `MAX_EVIDENCE_PER_SECTION` | `100` | Maximum sampled comment-anchor pairs per section. |
| `MAX_CHARS_ANCHOR` | `1000` | Character limit for each anchor excerpt. |
| `MAX_CHARS_COMMENT` | `1000` | Character limit for each reviewer comment. |
| `RANDOM_STATE` | `7` | Seed for evidence sampling. |

The request also sets `temperature=0.2` and `max_output_tokens=1400`. These are implementation defaults, not a guarantee of compatibility with every model configuration or enough output space for every response.

Inspect the `errors` dictionary and section counts after generation. A target of ten items is not a guarantee: unavailable sections, duplicate removal, or failed requests can reduce the output. The prompt asks the model to generalise when evidence is repetitive, so sparse sections require particular scrutiny.

## Prepare the final checklist

**Human revision is a required handoff.** Review each proposed requirement, fail signal, fix pattern, and priority against the underlying evidence before using it to assess new ToRs. Remove unsupported or redundant items and decide which checks apply to the intended evaluation.

The generated CSV contains:

```text
section, requirement, fail_signal, fix_pattern, priority, evidence_count
```

The review assistant expects the following **exact input headers**:

| Required column | Expected content |
| --- | --- |
| `ToR Section` | Section or overarching label used to contextualise the check. |
| `fail_signal_id` | A unique numeric identifier; use stable, non-missing integers. |
| `fail_signal` | Observable condition indicating a possible weakness in the ToR. |
| `fix_pattern` | Human-reviewed revision instruction. |
| `fix_mode (0 = instruction-only, 1 = instruction plus suggestion)` | Numeric `0` for instructions only, or `1` to permit a generated suggestion when the evidence gate passes. |
| `importance (0 = low, 1 = medium, 2 = high)` | Numeric `0`, `1`, or `2`. |
| `activated (1 = yes, 0 = no)` | Numeric `1` to include the check, or `0` to exclude it. |

Additional descriptive columns may be retained, but they do not replace the required headers. The loader converts the coded fix mode and importance values into text labels internally; the input CSV should contain the numeric codes above.

**Encoding matters:** generation writes UTF-8 CSVs, while the review notebook reads `Final_Checklist.csv` with `encoding="latin1"`. Match the loader to the file's actual encoding. For a UTF-8 final checklist, replace the existing load line with:

```python
checklist_df = pd.read_csv(CHECKLIST_PATH, encoding="utf-8")
```

<details>
<summary>Optional starter conversion: draft schema to review schema</summary>

Add this as an optional cell **after checklist generation**. It creates a candidate for human review, not an approved final checklist. It defaults all checks to active and instruction-only, and maps generated priorities to the required numeric codes.

```python
from pathlib import Path
import pandas as pd

draft = pd.read_csv(CHECKLIST_OUT_PATH, encoding="utf-8")
required = {"section", "requirement", "fail_signal", "fix_pattern", "priority"}
missing = required - set(draft.columns)
if missing:
    raise ValueError(f"Missing draft columns: {sorted(missing)}")
if draft.empty:
    raise ValueError("The draft checklist is empty.")

candidate = draft.rename(columns={
    "section": "ToR Section",
    "requirement": "Requirement",
}).copy()

candidate["fail_signal_id"] = range(1, len(candidate) + 1)
candidate["activated (1 = yes, 0 = no)"] = 1
candidate["fix_mode (0 = instruction-only, 1 = instruction plus suggestion)"] = 0

importance_col = "importance (0 = low, 1 = medium, 2 = high)"
candidate[importance_col] = candidate["priority"].map({
    "Low": 0, "Medium": 1, "High": 2,
})
if candidate[importance_col].isna().any():
    raise ValueError("Unexpected priority value; review the draft before conversion.")
candidate[importance_col] = candidate[importance_col].astype(int)

candidate_path = Path(CHECKLIST_OUT_DIR) / "Final_Checklist_candidate.csv"
if candidate_path.exists():
    raise FileExistsError(f"Refusing to overwrite: {candidate_path}")
candidate.to_csv(candidate_path, index=False, encoding="utf-8")
print(f"Candidate saved for human review: {candidate_path}")
```

Review the candidate, adjust activation and suggestion permissions, and save the approved version as `Final_Checklist.csv`. Keep identifiers stable after approval; regenerating IDs can change the meaning of ID-based activation settings. Use the UTF-8 loader shown above for this example.

</details>

## Review a draft ToR

Open `ToR_Review_Assistant.ipynb` and complete setup. Replace its configuration cell with your paths and document selection. For example:

```python
from pathlib import Path

PROJECT_ROOT = "/content/drive/MyDrive/ToR_Review_Project"
ENV_PATH = f"{PROJECT_ROOT}/.env"
CHECKLIST_PATH = f"{PROJECT_ROOT}/Checklist_Documents/Final_Checklist.csv"
INPUT_DIR = f"{PROJECT_ROOT}/Input"
TOR_FILES = {"my_draft": f"{INPUT_DIR}/draft_tor.docx"}
SELECTED_TOR = "my_draft"
TOR_INPUT_PATH = TOR_FILES[SELECTED_TOR]
LOGO_PATH = f"{INPUT_DIR}/Logo.jpg"
OUTPUT_DIR = f"{PROJECT_ROOT}/Output"

Path(OUTPUT_DIR).mkdir(parents=True, exist_ok=True)
```

The original document-selection dictionary points to files not supplied here. Replace those entries rather than relying on the example selection. The report cell unconditionally calls `add_picture(LOGO_PATH, ...)`: provide an image you are authorised to use, or remove/guard that line before generating an unbranded report.

### Select the checks

Load the reviewed checklist and inspect its columns. Replace or clear the notebook's hard-coded `ids_to_deactivate` list; it disables 28 specific IDs from the original checklist and should not be applied blindly to a new one.

Setting `ids_to_deactivate = []` prevents additional deactivation, but does not reactivate rows already disabled in the CSV or current session. To deliberately enable every item, use:

```python
checklist_df["activated (1 = yes, 0 = no)"] = 1
ids_to_deactivate = []
```

Run this before building `fail_signals_df`. After changing activation, rerun the downstream cells so the review uses the new selection. At least one active item is needed by the current implementation.

### Inspect parsing before making model calls

Run the parser, cleaning, and document-summary cells. Inspect the paragraph samples, section previews, `(unknown)` assignments, and detected-heading warnings against the original document. The assistant evaluates the **full parsed ToR for each active fail signal**, not just the item's named section.

After confirming the extracted content, run classification, guidance generation, summarisation, and export in sequence. These stages make external model requests. Review the final findings before sharing the report or applying any proposed wording.

## Outputs

| File | Contents |
| --- | --- |
| `ToR_comments_table_clean.csv` | `document`, `comment_id`, `bucket`, `heading_raw`, `anchor_paragraph`, and `comment`. Despite cleaning, this is not guaranteed to be fully anonymised. |
| `ToR_Checklist_Before_Human_Feedback.csv` | Draft requirements, fail signals, fix patterns, priorities, and estimated evidence counts. |
| `ToR_Review_Data.xlsx` | All active checks, including decisions, summaries where generated, rationales, target sections, instructions, suggested text, evidence paragraph IDs, and evidence excerpts. |
| `ToR_Review_Report.docx` | Review scope, aggregate decision counts, prioritised key issues, and detailed findings for triggered and uncertain checks, with internal links from key issues to details. |

The Excel export uses these columns:

```text
fail_signal_id, fail_signal, decision, summary_text, rationale,
target_section, instruction_text, suggested_text,
evidence_paragraph_ids, evidence_text
```

The Word report does **not** currently print the evidence excerpts or paragraph IDs. Use the Excel file for those references. Importance is used for Word-report grouping but is not included in the current Excel export columns.

### Interpreting decisions and guidance

| Decision | Interpretation |
| --- | --- |
| `TRIGGERED` | The model identified the fail signal; review the cited evidence and revision instruction. |
| `NOT_TRIGGERED` | The model did not identify that fail signal in the parsed content. This is not proof that the ToR is complete or correct. |
| `UNCERTAIN` | The assessment is inconclusive. This can also indicate that the classification response could not be parsed as JSON. |

Only triggered checks receive revision instructions. A text suggestion is attempted only when the checklist enables `instruction_plus_suggestion` and the evidence gate finds at least one referenced paragraph and at least 120 characters of evidence text. This is a simple threshold, not a semantic verification of sufficient evidence. A suggestion can still be empty when generation or JSON parsing fails, or when no suitable wording is returned.

Paragraph IDs refer to the parser's text entries, including table rows, rather than Word page numbers. `evidence_count` in the generated checklist is a scaled estimate based on model-selected sampled evidence IDs, not a verified count of distinct source documents. The draft CSV does not preserve the model's evidence-ID lists.

**Exports use fixed filenames.** Re-running the export cells replaces previous results. Copy or rename outputs between reviews. The review export cells build paths directly from `PROJECT_ROOT`; changing `OUTPUT_DIR` alone does not change their destinations.

## Models and request volume

These are the model identifiers configured in the notebooks, not claims about current availability or the newest models:

| Task | Default identifier | Configuration location |
| --- | --- | --- |
| Checklist generation | `gpt-5.2` | `MODEL` in `Checklist_Generation.ipynb`. |
| Fail-signal classification | `gpt-5.2` | Default argument in `check_fail_signal`. |
| Revision suggestions | `gpt-5.2` | Default argument in `generate_guidance`. |
| Finding summaries | `gpt-4o-mini` | Default argument in `generate_summary_item`. |
| Project-title extraction, when rules fail | `gpt-4o-mini` | Default argument in `detect_project_name_from_docx`. |

Model configuration is distributed across these functions. Update every relevant location when changing models, and verify support for the Responses API, structured outputs where used, and the supplied request parameters.

At the notebook-loop level, a review makes one classification request per active check, one suggestion request per eligible triggered check, one summary request per triggered or uncertain check, and optionally one project-title request. Checklist generation makes one initial request per represented section, with smaller-evidence retries in the generation loop. This excludes any SDK-level retries.

Classification and suggestion requests include the full parsed ToR, so document length and the number of active items affect request volume and token usage. The notebooks do not implement persistent response caching, automatic cost estimates, or resumable review execution. `CACHE_DIR` and `LOG_DIR` are defined in the original configuration but are not implemented caching or logging systems.

## Privacy and responsible use

> **This is not a fully local or fully anonymised workflow.** Selected evidence and ToR content are sent to the OpenAI API. In Colab, working files are also stored or processed in the configured Google environment.

The notebooks mask email addresses and detected multi-token person names in selected text fields. Checklist-generation prompts additionally mask URLs, years, and numbers of three or more digits. These measures are partial: single-token names, missed entities, organisations, locations, filenames, and other identifying details can remain. `heading_raw` is cleaned but is not passed through person-name anonymisation, and the review's heading-derived metadata is also not fully redacted.

**The project-title fallback is a separate exposure path.** It reads the original `.docx` and may send the first 6,000 characters of extracted, unredacted text to a model. For sensitive documents, remove or replace that cell with a manually supplied `project_title`, or redact the fallback input before using it.

Only process documents under an appropriate authorisation and data-handling arrangement. Inspect outgoing text, retain human oversight, and treat reports and intermediate datasets as sensitive. Model-generated wording must not be accepted as a factual source.

Before publishing, clear notebook outputs and inspect source cells, filenames, paths, and version history for confidential content. A suggested `.gitignore` for the working layout is:

```gitignore
.env
.env.*
!.env.example
.venv/
.ipynb_checkpoints/
__pycache__/
/Checklist_Documents/
/Input/
/Output/
/Cache/
/logs/
```

Adapt these patterns to your actual paths. Ignore rules do not remove files already tracked by Git, and do not remove sensitive text embedded in notebook outputs. Publish only deliberately sanitised examples and authorised branding.


## Reproducibility and quality checks

The supplied notebooks contain saved outputs, but no automated test suite, pinned environment, or independent accuracy benchmark. A saved checklist-generation run reports 35 source documents, 707 extracted comment-anchor pairs, 677 evidence rows after completeness filtering and deduplication, and 100 draft checklist items across ten groups. These are historical run outputs, not guaranteed results for new data.

The fixed sampling seed controls only evidence sampling under the same input ordering and data. The extraction loop does not consistently sort the input file list, model responses can vary, and model identifiers are not pinned to dated snapshots.

For an auditable run, record the notebook revision, dependency versions, model identifiers and settings, input versions, approved checklist version, active IDs, and review date. Preserve the human-reviewed checklist separately from generated candidates. Check extraction coverage, section assignments, generation errors, ID joins, evidence references, uncertain decisions, and both exported files before treating the run as complete.

## Contributing

When reporting a problem, identify the notebook and cell or function, describe expected and actual behaviour, and include a sanitised traceback and environment details. Do not attach confidential ToRs, credentials, or unredacted model responses.

Useful contributions include schema validation, empty-result handling, synthetic parser tests, consistent model configuration, dependency pinning, safer redaction, request checkpointing, and evidence excerpts in the Word report. Describe behavioural changes clearly and include a minimal non-sensitive example that demonstrates them.
