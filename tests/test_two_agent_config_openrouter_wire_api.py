"""Tests for two-agent Codex provider wire API selection."""

from __future__ import annotations

import json
import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from two_agent_framework import config as two_agent_config


def _write_provider_config(path: Path, **overrides: str) -> None:
    payload = {
        "provider_id": "openrouter",
        "provider_name": "OpenRouter",
        "base_url": "https://openrouter.ai/api/v1",
        "env_key": "OPENROUTER_API_KEY",
        "api_key": "sk-test",
        "model": "openai/gpt-5.4",
        "wire_api": "responses",
    }
    payload.update(overrides)
    path.write_text(json.dumps(payload), encoding="utf-8")


class OpenRouterWireApiTests(unittest.TestCase):
    def test_openrouter_uses_responses_wire_api_for_current_codex_cli(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            provider_path = Path(tmpdir) / "local_provider.json"
            _write_provider_config(provider_path)

            with patch.object(two_agent_config, "LOCAL_PROVIDER_CONFIG", provider_path), patch.dict(
                os.environ,
                {},
                clear=True,
            ):
                overrides = two_agent_config._codex_provider_overrides()

        self.assertIn("model_providers.openrouter.wire_api=\"responses\"", overrides)
        self.assertNotIn("model_providers.openrouter.wire_api=\"chat\"", overrides)

    def test_openrouter_defaults_to_responses_when_wire_api_is_omitted(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            provider_path = Path(tmpdir) / "local_provider.json"
            _write_provider_config(provider_path, wire_api="")

            with patch.object(two_agent_config, "LOCAL_PROVIDER_CONFIG", provider_path), patch.dict(
                os.environ,
                {},
                clear=True,
            ):
                overrides = two_agent_config._codex_provider_overrides()

        self.assertIn("model_providers.openrouter.wire_api=\"responses\"", overrides)

    def test_local_provider_can_be_disabled_for_direct_openai_comparison(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            provider_path = Path(tmpdir) / "local_provider.json"
            _write_provider_config(provider_path)

            with patch.object(two_agent_config, "LOCAL_PROVIDER_CONFIG", provider_path), patch.dict(
                os.environ,
                {"TWOWORK_CODEX_DISABLE_LOCAL_PROVIDER": "1"},
                clear=True,
            ):
                overrides = two_agent_config._codex_provider_overrides()
                model = two_agent_config.effective_codex_model()

        self.assertEqual(overrides, [])
        self.assertIsNone(model)

    def test_non_openrouter_provider_preserves_configured_wire_api(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            provider_path = Path(tmpdir) / "local_provider.json"
            _write_provider_config(
                provider_path,
                provider_id="custom_responses",
                provider_name="Custom Responses",
                wire_api="responses",
            )

            with patch.object(two_agent_config, "LOCAL_PROVIDER_CONFIG", provider_path), patch.dict(
                os.environ,
                {},
                clear=True,
            ):
                overrides = two_agent_config._codex_provider_overrides()

        self.assertIn("model_provider=\"custom_responses\"", overrides)
        self.assertIn("model_providers.custom_responses.wire_api=\"responses\"", overrides)


if __name__ == "__main__":
    unittest.main()
