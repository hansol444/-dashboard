"""
integrations.py — SharePoint 업로드 & Slack 알림
번역 완료 후 자동 실행되는 외부 연동 모듈

사전 설정 필요:
  1. SharePoint: Azure App Registration → Client ID / Secret / Tenant ID
  2. Slack: Bot Token (xoxb-...) → Slack App에 chat:write 권한 필요
"""

import base64
from pathlib import Path

import config


# ─── SharePoint 업로드 ─────────────────────────────────────────────────────────

def _get_ms_token() -> str:
    """Microsoft Graph API 액세스 토큰 획득 (App 인증)."""
    import msal

    app = msal.ConfidentialClientApplication(
        client_id=config.SHAREPOINT_CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{config.SHAREPOINT_TENANT_ID}",
        client_credential=config.SHAREPOINT_CLIENT_SECRET,
    )
    result = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )
    if "access_token" not in result:
        raise RuntimeError(
            f"Microsoft 토큰 획득 실패: {result.get('error_description', result)}"
        )
    return result["access_token"]


def _share_url_to_graph_id(share_url: str) -> str:
    """SharePoint 공유 링크를 Graph API share ID 형식으로 변환."""
    encoded = (
        base64.b64encode(share_url.encode())
        .decode()
        .rstrip("=")
        .replace("+", "-")
        .replace("/", "_")
    )
    return "u!" + encoded


def upload_to_sharepoint(local_file_path: str) -> tuple[bool, str]:
    """
    번역된 PPT를 SharePoint 공유 폴더에 업로드.
    반환: (성공 여부, 메시지)
    """
    if not all([
        config.SHAREPOINT_CLIENT_ID,
        config.SHAREPOINT_CLIENT_SECRET,
        config.SHAREPOINT_TENANT_ID,
    ]):
        return False, "SharePoint 자격증명이 설정되지 않았습니다 (config.py 확인)"

    try:
        import requests

        token = _get_ms_token()
        filename = Path(local_file_path).name
        share_id = _share_url_to_graph_id(config.SHAREPOINT_SHARE_URL)

        upload_url = (
            f"https://graph.microsoft.com/v1.0"
            f"/shares/{share_id}/driveItem:/{filename}:/content"
        )
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/octet-stream",
        }

        with open(local_file_path, "rb") as f:
            resp = requests.put(upload_url, headers=headers, data=f, timeout=60)

        if resp.status_code in (200, 201):
            return True, f"{filename} 업로드 완료"
        else:
            return False, f"업로드 실패 (HTTP {resp.status_code}): {resp.text[:200]}"

    except ImportError:
        return False, "msal 또는 requests 패키지가 필요합니다: pip install msal requests"
    except Exception as e:
        return False, f"SharePoint 오류: {str(e)}"


# ─── Slack 알림 ────────────────────────────────────────────────────────────────

def send_slack_notification(filename: str, overflow_count: int = 0) -> tuple[bool, str]:
    """
    번역 완료 후 Slack 채널에 메시지 전송.
    반환: (성공 여부, 메시지)
    """
    if not config.SLACK_BOT_TOKEN:
        return False, "SLACK_BOT_TOKEN이 설정되지 않았습니다 (config.py 확인)"

    try:
        from slack_sdk import WebClient
        from slack_sdk.errors import SlackApiError

        client = WebClient(token=config.SLACK_BOT_TOKEN)

        overflow_note = f" (글자 초과 {overflow_count}개 - 빨간 항목 확인 필요)" if overflow_count > 0 else ""
        text = f"{config.SLACK_MESSAGE_TEMPLATE} — `{filename}`{overflow_note}"

        client.chat_postMessage(
            channel=config.SLACK_CHANNEL_ID,
            text=text,
        )
        return True, "Slack 메시지 전송 완료"

    except ImportError:
        return False, "slack_sdk 패키지가 필요합니다: pip install slack_sdk"
    except Exception as e:
        return False, f"Slack 오류: {str(e)}"
