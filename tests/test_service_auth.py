from __future__ import annotations

import json
import unittest
from unittest.mock import patch

from fastapi import HTTPException
from starlette.requests import Request

import src.service.app as service


def _request(headers: dict[str, str]) -> Request:
    return Request(
        {
            "type": "http",
            "headers": [(key.lower().encode(), value.encode()) for key, value in headers.items()],
        }
    )


class _AuthResponse:
    def __init__(self, payload: dict[str, str]) -> None:
        self._payload = json.dumps(payload).encode("utf-8")

    def __enter__(self) -> _AuthResponse:
        return self

    def __exit__(self, *_: object) -> None:
        return None

    def read(self) -> bytes:
        return self._payload


class ServiceAuthTest(unittest.TestCase):
    def setUp(self) -> None:
        with service._SUPABASE_AUTH_CACHE_LOCK:
            service._SUPABASE_AUTH_CACHE.clear()

    def test_static_service_token_remains_available_for_server_calls(self) -> None:
        with patch.object(service.CONFIG, "auth_token", "service-secret"), patch.object(
            service.CONFIG, "supabase_url", ""), patch.object(
            service.CONFIG, "supabase_anon_key", ""
        ):
            identity = service._require_auth(_request({"Authorization": "Bearer service-secret"}))

        self.assertTrue(identity.is_service)
        self.assertIsNone(identity.user_id)

    def test_supabase_access_token_resolves_to_user_identity(self) -> None:
        with patch.object(service.CONFIG, "auth_token", "service-secret"), patch.object(
            service.CONFIG, "supabase_url", "https://example.supabase.co"), patch.object(
            service.CONFIG, "supabase_anon_key", "anon-key"), patch.object(
            service, "urlopen", return_value=_AuthResponse({"id": "user-1"})
        ) as urlopen:
            identity = service._require_auth(_request({"Authorization": "Bearer user-jwt"}))

        self.assertFalse(identity.is_service)
        self.assertEqual(identity.user_id, "user-1")
        urlopen.assert_called_once()

    def test_missing_credentials_are_rejected(self) -> None:
        with patch.object(service.CONFIG, "auth_token", "service-secret"):
            with self.assertRaises(HTTPException) as context:
                service._require_auth(_request({}))

        self.assertEqual(context.exception.status_code, 401)


if __name__ == "__main__":
    unittest.main()
