"""统一 AI 客户端抽象层（v7）。"""

from __future__ import annotations

import time
from dataclasses import dataclass
from enum import Enum
from typing import Any, Dict, List, Optional, Tuple

from openai import (
    APIConnectionError,
    APIStatusError,
    APITimeoutError,
    AuthenticationError,
    OpenAI,
    RateLimitError,
)


class AIErrorType(Enum):
    NETWORK_CONNECTION = "network_connection"
    AUTHENTICATION = "authentication"
    RATE_LIMIT = "rate_limit"
    TIMEOUT = "timeout"
    API_ERROR = "api_error"
    UNKNOWN = "unknown"


@dataclass
class AIError:
    type: AIErrorType
    message: str
    solution: str = ""


@dataclass
class PlatformConfig:
    max_concurrency: int = 1
    timeout: float = 120.0
    max_retries: int = 2
    retry_delay: float = 1.5


PLATFORM_CONFIGS: Dict[str, PlatformConfig] = {
    "硅基流动": PlatformConfig(max_concurrency=3, timeout=180.0),
    "DeepSeek": PlatformConfig(max_concurrency=3, timeout=120.0),
    "OpenAI": PlatformConfig(max_concurrency=2, timeout=120.0),
    "通义千问(阿里云百炼)": PlatformConfig(max_concurrency=3, timeout=120.0),
    "智谱AI": PlatformConfig(max_concurrency=2, timeout=120.0),
    "百度千帆": PlatformConfig(max_concurrency=2, timeout=120.0),
    "火山方舟": PlatformConfig(max_concurrency=2, timeout=120.0),
    "腾讯混元": PlatformConfig(max_concurrency=2, timeout=120.0),
    "Groq": PlatformConfig(max_concurrency=5, timeout=60.0),
    "自定义/本地": PlatformConfig(max_concurrency=1, timeout=300.0, max_retries=1),
}

MODEL_SPECIFIC_CONFIGS: Dict[str, PlatformConfig] = {}


class ProgressManager:
    """线程安全的简易进度缓存（供后续扩展）。"""

    def __init__(self) -> None:
        self._steps: Dict[str, float] = {}

    def update(self, step: str, progress: float) -> None:
        self._steps[step] = max(0.0, min(1.0, progress))

    def get(self, step: str) -> float:
        return self._steps.get(step, 0.0)

    def reset(self) -> None:
        self._steps.clear()


progress_manager = ProgressManager()


class AIClient:
    def __init__(self, platform: str, base_url: str, api_key: str, model: str) -> None:
        self.platform = platform or ""
        self.base_url = (base_url or "").rstrip("/")
        self.api_key = api_key or ""
        self.model = model or ""

        config = self._resolve_config()
        self.timeout = config.timeout
        self.max_retries = config.max_retries
        self.retry_delay = config.retry_delay

        client_kwargs: Dict[str, Any] = {
            "api_key": self.api_key,
            "timeout": self.timeout,
            "max_retries": 0,
        }
        if self.base_url:
            client_kwargs["base_url"] = self.base_url
        self._raw_client = OpenAI(**client_kwargs)

    @property
    def client(self) -> OpenAI:
        return self._raw_client

    def _resolve_config(self) -> PlatformConfig:
        if self.model and self.model in MODEL_SPECIFIC_CONFIGS:
            return MODEL_SPECIFIC_CONFIGS[self.model]
        return PLATFORM_CONFIGS.get(self.platform, PlatformConfig())

    def _classify_error(self, exc: Exception) -> AIError:
        if isinstance(exc, APIConnectionError):
            return AIError(
                AIErrorType.NETWORK_CONNECTION,
                "网络连接失败",
                "请检查网络、代理设置及 API 地址是否正确",
            )
        if isinstance(exc, APITimeoutError):
            return AIError(
                AIErrorType.TIMEOUT,
                "请求超时",
                "请稍后重试，或切换到响应更快的模型",
            )
        if isinstance(exc, AuthenticationError):
            return AIError(
                AIErrorType.AUTHENTICATION,
                "API Key 认证失败",
                "请检查 Key 是否有效、是否已过期、是否与所选平台匹配",
            )
        if isinstance(exc, RateLimitError):
            return AIError(
                AIErrorType.RATE_LIMIT,
                "请求频率超限",
                "请降低并发数或稍后重试",
            )
        if isinstance(exc, APIStatusError):
            detail = getattr(exc, "message", None) or str(exc)
            return AIError(
                AIErrorType.API_ERROR,
                f"API 返回错误: {detail}",
                "请检查模型名称、配额及平台服务状态",
            )
        return AIError(
            AIErrorType.UNKNOWN,
            str(exc) or "未知错误",
            "请查看日志或联系平台支持",
        )

    def _extract_content(self, response: Any) -> Optional[str]:
        if not response or not getattr(response, "choices", None):
            return None
        choice = response.choices[0]
        message = getattr(choice, "message", None)
        if not message:
            return None
        content = getattr(message, "content", None)
        return content.strip() if isinstance(content, str) and content.strip() else None

    def generate(
        self,
        messages: List[Dict[str, str]],
        max_tokens: int = 2048,
        temperature: float = 0.3,
    ) -> Tuple[Optional[str], Optional[AIError]]:
        last_error: Optional[AIError] = None
        attempts = max(1, self.max_retries + 1)

        for attempt in range(attempts):
            try:
                response = self._raw_client.chat.completions.create(
                    model=self.model,
                    messages=messages,
                    max_tokens=max_tokens,
                    temperature=temperature,
                )
                content = self._extract_content(response)
                if content:
                    return content, None
                last_error = AIError(
                    AIErrorType.API_ERROR,
                    "模型返回空内容",
                    "请重试或更换模型",
                )
            except Exception as exc:
                last_error = self._classify_error(exc)
                if last_error.type not in (
                    AIErrorType.NETWORK_CONNECTION,
                    AIErrorType.TIMEOUT,
                    AIErrorType.RATE_LIMIT,
                ):
                    break
            if attempt < attempts - 1:
                time.sleep(self.retry_delay * (attempt + 1))

        return None, last_error

    def test_connection(self) -> Tuple[bool, str]:
        try:
            response = self._raw_client.chat.completions.create(
                model=self.model,
                messages=[{"role": "user", "content": "ping"}],
                max_tokens=5,
                temperature=0,
            )
            if self._extract_content(response):
                return True, "连接成功"
            return False, "连接失败：模型返回空响应"
        except Exception as exc:
            err = self._classify_error(exc)
            msg = err.message
            if err.solution:
                msg = f"{err.message}。{err.solution}"
            return False, msg


class AIClientFactory:
    @staticmethod
    def create_client(platform: str, base_url: str, api_key: str, model: str) -> AIClient:
        return AIClient(platform, base_url, api_key, model)
