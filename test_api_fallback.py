# -*- coding: utf-8 -*-

import io
import json
import unittest
import urllib.error
from unittest.mock import patch

from generate_bug_report import call_openai_compatible_chat


class _FakeResponse:
    def __init__(self, payload):
        self._payload = payload

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        return False

    def read(self):
        return self._payload


def _http_error(endpoint, status, detail):
    return urllib.error.HTTPError(
        endpoint,
        status,
        "test error",
        {},
        io.BytesIO(detail.encode("utf-8")),
    )


class ApiFallbackTests(unittest.TestCase):
    def test_cloudflare_1010_falls_back_to_chat_completions(self):
        calls = []
        user_agents = []

        def fake_urlopen(request, timeout):
            calls.append(request.full_url)
            user_agents.append(request.get_header("User-agent"))
            if request.full_url.endswith("/v1/responses"):
                raise _http_error(request.full_url, 403, "error code: 1010")
            return _FakeResponse(
                json.dumps({"choices": [{"message": {"content": "ok"}}]}).encode()
            )

        with patch("urllib.request.urlopen", side_effect=fake_urlopen):
            result = call_openai_compatible_chat(
                api_base_url="https://example.test",
                api_key="test-key",
                model_name="test-model",
                messages=[{"role": "user", "content": "hello"}],
                api_endpoint="/v1/responses",
            )

        self.assertEqual(result, "ok")
        self.assertEqual(
            calls,
            [
                "https://example.test/v1/responses",
                "https://example.test/v1/chat/completions",
            ],
        )
        self.assertTrue(all(user_agents))

    def test_401_does_not_fallback(self):
        calls = []

        def fake_urlopen(request, timeout):
            calls.append(request.full_url)
            raise _http_error(request.full_url, 401, '{"error":"Invalid token"}')

        with patch("urllib.request.urlopen", side_effect=fake_urlopen):
            with self.assertRaisesRegex(RuntimeError, "401"):
                call_openai_compatible_chat(
                    api_base_url="https://example.test",
                    api_key="test-key",
                    model_name="test-model",
                    messages=[{"role": "user", "content": "hello"}],
                    api_endpoint="/v1/responses",
                )

        self.assertEqual(calls, ["https://example.test/v1/responses"])


if __name__ == "__main__":
    unittest.main()
