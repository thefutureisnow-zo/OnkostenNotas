"""Tests voor gmail_client.py -- OAuth2 token handling."""

from unittest.mock import MagicMock, patch

import pytest
from google.auth.exceptions import RefreshError

from gmail_client import get_gmail_service


@pytest.fixture
def tmp_token(tmp_path):
    """Maak een dummy token.json aan."""
    token = tmp_path / "token.json"
    token.write_text('{"token": "x", "refresh_token": "r", "client_id": "c", "client_secret": "s"}')
    return token


@pytest.fixture
def tmp_secret(tmp_path):
    """Maak een dummy client_secret.json aan."""
    secret = tmp_path / "client_secret.json"
    secret.write_text('{"installed": {"client_id": "c", "client_secret": "s"}}')
    return secret


@patch("gmail_client.build")
@patch("gmail_client.InstalledAppFlow")
@patch("gmail_client.Credentials")
def test_refresh_error_falls_through_to_reauth(
    mock_creds_cls, mock_flow_cls, mock_build, tmp_token, tmp_secret, capsys
):
    """Bij RefreshError moet de flow opnieuw inloggen via de browser."""
    # Stel een verlopen credential in die RefreshError gooit
    mock_creds = MagicMock()
    mock_creds.valid = False
    mock_creds.expired = True
    mock_creds.refresh_token = "some_token"
    mock_creds.refresh.side_effect = RefreshError("invalid_grant")
    mock_creds_cls.from_authorized_user_file.return_value = mock_creds

    # De flow moet nieuwe creds teruggeven
    mock_new_creds = MagicMock()
    mock_new_creds.to_json.return_value = '{"token": "new"}'
    mock_flow = MagicMock()
    mock_flow.run_local_server.return_value = mock_new_creds
    mock_flow_cls.from_client_secrets_file.return_value = mock_flow

    get_gmail_service(tmp_secret, tmp_token)

    # Verify: refresh was attempted, then re-auth flow was used
    mock_creds.refresh.assert_called_once()
    mock_flow_cls.from_client_secrets_file.assert_called_once()
    mock_flow.run_local_server.assert_called_once()

    # Verify user-friendly message was printed
    captured = capsys.readouterr()
    assert "opnieuw inloggen" in captured.out

    # Verify new token was persisted to disk
    assert tmp_token.read_text(encoding="utf-8") == '{"token": "new"}'


@patch("gmail_client.build")
@patch("gmail_client.InstalledAppFlow")
@patch("gmail_client.Credentials")
def test_retryable_refresh_error_is_reraised(
    mock_creds_cls, mock_flow_cls, mock_build, tmp_token, tmp_secret
):
    """Tijdelijke RefreshError (5xx) moet opnieuw gegooid worden."""
    mock_creds = MagicMock()
    mock_creds.valid = False
    mock_creds.expired = True
    mock_creds.refresh_token = "some_token"
    class RetryableRefreshError(RefreshError):
        retryable = True
    mock_creds.refresh.side_effect = RetryableRefreshError("server_error")
    mock_creds_cls.from_authorized_user_file.return_value = mock_creds

    with pytest.raises(RefreshError):
        get_gmail_service(tmp_secret, tmp_token)

    # Re-auth flow should NOT have been triggered
    mock_flow_cls.from_client_secrets_file.assert_not_called()
