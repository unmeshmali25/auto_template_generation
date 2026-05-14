"""
OpenRouter and Custom LLM client adapters.
Supports chat completions with structured JSON output.
"""
import os
import json
import requests
from typing import Dict, Any, Optional


class OpenRouterClient:
    """Client for OpenRouter API chat completions."""

    def __init__(self, api_key: Optional[str] = None, base_url: str = "https://openrouter.ai/api"):
        self.api_key = api_key or os.getenv("OPENROUTER_API_KEY")
        if not self.api_key:
            raise ValueError("OPENROUTER_API_KEY not provided and not found in environment")
        self.base_url = base_url.rstrip("/")
        self.headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
            "HTTP-Referer": "https://localhost",  # required by OpenRouter
            "X-Title": "Communications Recommendations Generator",
        }

    def chat_completion(
        self,
        messages: list,
        model: str = "openai/gpt-4o-mini",
        temperature: float = 0.2,
        max_tokens: int = 4000,
        response_format: Optional[Dict[str, Any]] = None,
    ) -> Dict[str, Any]:
        """Send a chat completion request to OpenRouter."""
        url = f"{self.base_url}/v1/chat/completions"
        payload = {
            "model": model,
            "messages": messages,
            "temperature": temperature,
            "max_tokens": max_tokens,
        }
        if response_format:
            payload["response_format"] = response_format

        resp = requests.post(url, headers=self.headers, json=payload, timeout=120)
        resp.raise_for_status()
        data = resp.json()
        return data

    def extract_structured_json(
        self,
        system_prompt: str,
        user_prompt: str,
        model: str = "openai/gpt-4o-mini",
    ) -> Dict[str, Any]:
        """Call LLM with structured JSON response format and parse result."""
        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ]
        response_format = {"type": "json_object"}
        data = self.chat_completion(
            messages=messages,
            model=model,
            response_format=response_format,
        )
        content = data["choices"][0]["message"]["content"]
        return json.loads(content)


class CustomChatClient:
    """Client for custom generative AI chat API (e.g., internal enterprise endpoint)."""

    def __init__(
        self,
        client_key: Optional[str] = None,
        pass_key: Optional[str] = None,
        endpoint_url: Optional[str] = None,
        user_email: Optional[str] = None,
    ):
        self.client_key = client_key or os.getenv("YOUR_CLIENT_KEY")
        self.pass_key = pass_key or os.getenv("YOUR_PASS_KEY")
        self.endpoint_url = (endpoint_url or os.getenv("ENDPOINT_URL", "")).rstrip("/")
        self.user_email = user_email or os.getenv("YOUR_EMAIL")

        if not all([self.client_key, self.pass_key, self.endpoint_url]):
            raise ValueError(
                "CustomChatClient requires YOUR_CLIENT_KEY, YOUR_PASS_KEY, and ENDPOINT_URL. "
                "Set them via arguments or environment variables."
            )

        self.headers = {
            "x-generative-ai-client": self.client_key,
            "x-openapi-token": self.pass_key,
            "x-generative-ai-user-email": self.user_email,
            "Content-Type": "application/json",
        }

    def list_models(self) -> Dict[str, Any]:
        """List available models via GET request."""
        api_endpoint_url = f"{self.endpoint_url}/openapi/chat/v1/models"
        response = requests.get(
            api_endpoint_url,
            headers=self.headers,
            timeout=120,
        )
        response.raise_for_status()
        return response.json()

    def chat_completion(
        self,
        messages: list,
        model: str,
        temperature: float = 0.2,
        max_tokens: int = 4000,
        response_format: Optional[Dict[str, Any]] = None,
        **kwargs,
    ) -> Dict[str, Any]:
        """Send a chat completion request via POST to the custom endpoint."""
        url = f"{self.endpoint_url}/openapi/chat/v1/chat/completions"
        payload = {
            "model": model,
            "messages": messages,
            "temperature": temperature,
            "max_tokens": max_tokens,
        }
        if response_format:
            payload["response_format"] = response_format
        payload.update(kwargs)

        resp = requests.post(url, headers=self.headers, json=payload, timeout=120)
        resp.raise_for_status()
        return resp.json()

    def extract_structured_json(
        self,
        system_prompt: str,
        user_prompt: str,
        model: str,
    ) -> Dict[str, Any]:
        """Call LLM with structured JSON response format and parse result."""
        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ]
        response_format = {"type": "json_object"}
        data = self.chat_completion(
            messages=messages,
            model=model,
            response_format=response_format,
        )
        content = data["choices"][0]["message"]["content"]
        return json.loads(content)
