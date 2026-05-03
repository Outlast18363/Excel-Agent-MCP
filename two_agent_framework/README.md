# Two-Agent Framework

This package runs spreadsheet tasks with a compact agent loop:

```text
Worker -> Evaluator
```

The `Distiller` is used only when the Evaluator asks for a reset:

```text
Worker -> Evaluator -> Distiller -> restore snapshot -> Worker
```

The framework intentionally removes planner/executor handoff state and does
not create `plan.md`. The Worker owns discovery and implementation, writes
`handover/impl_report.md`, and the Evaluator verifies from `snapshots/`,
`final_result/`, the task text, and the implementation report.

Run it with the same CLI shape as the multi-agent runner:

```bash
python -m two_agent_framework.runner \
  --task "Update the workbook." \
  --workbook-dir C:\path\to\run\workbook \
  --run-dir C:\path\to\run \
  --task-id 84
```

## Using OpenRouter

The two-agent framework does not patch Codex internals or rewrite API keys. It
uses the normal Codex CLI runtime override path: the framework adds `-c`
configuration arguments to each `codex exec` subprocess.

OpenRouter's OpenAI-compatible base URL is:

```text
https://openrouter.ai/api/v1
```

Create a local provider file at `two_agent_framework/local_provider.json`:

```json
{
  "provider_id": "openrouter",
  "provider_name": "OpenRouter",
  "base_url": "https://openrouter.ai/api/v1",
  "env_key": "OPENROUTER_API_KEY",
  "api_key": "sk-or-v1-...",
  "model": "openai/gpt-5.4",
  "reasoning_effort": "high",
  "wire_api": "responses"
}
```

This file is ignored by git. The framework reads it at runtime and injects
`api_key` only into the `codex exec` child process environment, so the Codex app
can keep using its normal OpenAI login/quota.

Then run the framework normally:

```powershell
python -m two_agent_framework.runner `
  --task "Update the workbook." `
  --workbook-dir C:\path\to\run\workbook `
  --run-dir C:\path\to\run `
  --task-id 84
```

Choose `model` from the OpenRouter model list. OpenRouter model IDs include the
provider prefix, such as `openai/gpt-5.4` or
`anthropic/claude-sonnet-4.6`.

The resulting Codex command includes runtime overrides equivalent to:

```text
-c model_provider="openrouter"
-c model_providers.openrouter.name="OpenRouter"
-c model_providers.openrouter.base_url="https://openrouter.ai/api/v1"
-c model_providers.openrouter.env_key="OPENROUTER_API_KEY"
-c model_providers.openrouter.wire_api="responses"
-c model_reasoning_effort=high
```

Codex CLI 0.118+ no longer accepts `wire_api="chat"` for custom model
providers, so OpenRouter must be configured with `wire_api="responses"`.
OpenRouter's Responses API is currently a beta compatibility layer; with
`codex exec --json`, that beta path can finish a turn after a normal assistant
message instead of continuing into the shell/file tool call needed to create
`handover/impl_report.md`.

For production Worker runs that require reliable shell/file tool continuation,
use Codex's direct OpenAI/ChatGPT provider instead of OpenRouter. You can
temporarily ignore `local_provider.json` for a direct-provider comparison with:

```powershell
$env:TWOWORK_CODEX_DISABLE_LOCAL_PROVIDER = "1"
```

Unset that variable to return to the local OpenRouter config.

If a run fails before the Worker starts, check that `local_provider.json`
contains a valid API key and that the selected model supports the Codex features
you need, especially tool calling and streaming.

## Minimal implementation-report probe

The framework includes an isolated probe that uses the same Worker Codex config
and asks only for `handover/impl_report.md` to be created:

```powershell
python -m two_agent_framework.minimal_impl_write_probe --disable-excel-mcp
```

The two-agent baseline also has a wrapper testcase:

```powershell
py -3 ..\two_agent_baseline\test_cases\minimal_impl_report_probe_case.py --disable-excel-mcp
```

The probe fails if Codex exits cleanly but the report file is missing, or if the
JSON trace contains no command/file tool attempt.

To compare direct OpenAI/ChatGPT auth against the local OpenRouter config:

```powershell
py -3 ..\two_agent_baseline\test_cases\minimal_impl_report_probe_case.py --disable-excel-mcp --direct-openai --model gpt-5.4
```
