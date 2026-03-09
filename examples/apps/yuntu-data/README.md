# opt

`opt` implements a two-phase strategy:

1. `PHASE_A` (agent-first):
   - Agent performs natural-language navigation and opens target report.
   - Generic browser actions are enabled.
2. `PHASE_B` (controlled execution):
   - Agent uses module tools only for extraction/download.
   - Generic browser actions are disabled.

## Key files

- `pipeline.py`: two-phase orchestration (`optTaskPipeline`)
- `prompt_builder.py`: short prompts for phase A/B
- `dsl.py`: compact route DSL parser
- `specs.py`: DSL overrides on top of base profile specs

## DSL example

```text
人群分析 > 人群画像分析 > 画像分析 > [select]5A人群资产=投后 > [select]触点=全部 > [click]5A人群总资产 > [download]audience_profile_5a_asset
```

