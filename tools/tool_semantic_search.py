from collections.abc import Generator
from typing import Any

import json
import requests

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage


class ToolSemanticSearch(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        TENANT_ID = self.runtime.credentials["tenant_id"]
        CLIENT_ID = self.runtime.credentials["client_id"]
        CLIENT_SECRET = self.runtime.credentials["client_secret"]
        CONTAINER_ID = self.runtime.credentials["container_id"]
        search_query = tool_parameters["search_query"]
        access_token = self.authenticate(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
        if not access_token:
            yield self.create_json_message(
                {
                    "search_result": None,
                    "message": "get access_token failed",
                }
            )
            return
        search_result = self.semantic_search(access_token, search_query)
        if not search_result:
            yield self.create_json_message(
                {
                    "search_result": search_result,                 
                    "message": "get search_result failed",
                }
            )
            return
        yield self.create_json_message(
            {
                "search_result": search_result,                 
                "message": "get search_result success",
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

    # 语义搜索
    def semantic_search(self, access_token, search_query):
        """
        在 SharePoint Embedded 中进行语义搜索
        参考文档: https://learn.microsoft.com/zh-cn/sharepoint/dev/embedded/development/content-experiences/search-content

        Args:
            access_token: 访问令牌
            search_query: 搜索查询文本
            container_type_id: 容器类型 ID（可选）

        Returns:
            str: 搜索结果的 JSON 字符串，失败则返回 None
        """
        print(f"正在进行语义搜索: {search_query}...")
        try:
            # 调用 Microsoft Graph Search API
            url = "https://graph.microsoft.com/v1.0/search/query"
            headers = {
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/json",
            }

            # 构建查询字符串
            query_string = search_query
            payload = {
                "requests": [
                    {
                        "entityTypes": ["driveItem"],
                        "query": {"queryString": query_string},
                        "sharePointOneDriveOptions": {"includeHiddenContent": True},
                        "region": "APC",  # 亚太区域，根据错误信息调整
                        "from": 0,
                        "size": 25,
                        "fields": [
                            "id",
                            "name",
                            "createdDateTime",
                            "lastModifiedDateTime",
                            "contentPreview",
                        ],
                    }
                ]
            }
            response = requests.post(url, headers=headers, json=payload)
            response.raise_for_status()
            if response.status_code == 200:
                results = response.json()
                # 清洗JSON数据
                cleaned_results = []
                for item in results.get("value", []):
                    for container in item.get("hitsContainers", []):
                        for hit in container.get("hits", []):
                            # 从hit对象或resource对象中获取name字段
                            name = hit.get("name")
                            if not name and "resource" in hit:
                                name = hit["resource"].get("name")
                            cleaned_hit = {
                                "rank": hit.get("rank"),
                                "summary": hit.get("summary"),
                                "name": name,
                            }
                            cleaned_results.append(cleaned_hit)
                return json.dumps(cleaned_results, ensure_ascii=False)
            else:
                print(f"搜索失败: {response.status_code} - {response.text}")
                return None
        except Exception as e:
            print(f"搜索失败: {e}")
            return None
