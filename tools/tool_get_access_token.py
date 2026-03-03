from collections.abc import Generator
from typing import Any

import requests

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage


class ToolGetAccessToken(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        TENANT_ID = self.runtime.credentials["tenant_id"]
        CLIENT_ID = self.runtime.credentials["client_id"]
        CLIENT_SECRET = self.runtime.credentials["client_secret"]
        CONTAINER_ID = self.runtime.credentials["container_id"]

        access_token = self.authenticate(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
        if not access_token:
            yield self.create_json_message(
                {
                    "access_token": access_token,
                    "container_id": CONTAINER_ID,
                    "message": "get access_token failed",
                }
            )
            return
        yield self.create_json_message(
            {
                "access_token": access_token,
                "container_id": CONTAINER_ID,
                "message": "get access_token success",
            }
        )

    def authenticate(self, TENANT_ID: str, CLIENT_ID: str, CLIENT_SECRET: str):
        """
        使用 Client Credential 流程获取访问令牌
        """
        print("正在进行身份验证...")
        try:
            # 获取访问令牌
            token_url = (
                f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
            )
            token_data = {
                "grant_type": "client_credentials",
                "client_id": CLIENT_ID,
                "client_secret": CLIENT_SECRET,
                "scope": "https://graph.microsoft.com/.default",
            }
            token_response = requests.post(token_url, data=token_data)
            token_response.raise_for_status()
            token_json = token_response.json()
            access_token = token_json["access_token"]
            print("身份验证成功！")
            return access_token
        except Exception as e:
            print(f"身份验证失败：{e}")
            return None
