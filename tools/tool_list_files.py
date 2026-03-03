from collections.abc import Generator
from typing import Any

import json
import requests

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage


class ToolListFile(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        TENANT_ID = self.runtime.credentials["tenant_id"]
        CLIENT_ID = self.runtime.credentials["client_id"]
        CLIENT_SECRET = self.runtime.credentials["client_secret"]
        CONTAINER_ID = self.runtime.credentials["container_id"]

        access_token = self.authenticate(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
        if not access_token:
            yield self.create_json_message(
                {
                    "files_list": None,
                    "message": "get access_token failed",
                }
            )
            return
        files_list = self.list_files(access_token, CONTAINER_ID)
        if not files_list:
            yield self.create_json_message(
                {
                    "files_list": None,
                    "message": "get files_list failed",
                }
            )
            return
        yield self.create_json_message(
            {
                "files_list": files_list,
                "message": "get files_list success",
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

    # 列出文件
    def list_files(self, access_token, container_id):
        """
        列出容器中的文件
        """
        print("正在列出容器中的文件...")
        try:
            # 调用 Microsoft Graph API 获取文件列表
            # 使用 Microsoft Graph API drives 端点格式
            url = (
                f"https://graph.microsoft.com/v1.0/drives/{container_id}/root/children"
            )
            headers = {"Authorization": f"Bearer {access_token}"}
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            if response.status_code == 200:
                items = response.json()
                # 提取value数组并只保留@microsoft.graph.downloadUrl和name字段
                cleaned_files = []
                for item in items.get("value", []):
                    cleaned_item = {
                        "@microsoft.graph.downloadUrl": item.get(
                            "@microsoft.graph.downloadUrl"
                        ),
                        "name": item.get("name"),
                    }
                    cleaned_files.append(cleaned_item)
                return json.dumps(cleaned_files, ensure_ascii=False)
            else:
                return None
        except Exception as e:
            return e.json()
