"""
OpenRouter LLM client adapter.
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
