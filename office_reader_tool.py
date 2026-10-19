"""
title: Office Document Reader
description: Converts Word, Excel, PowerPoint and PDF files to Markdown text. Takes either a pre-authenticated download URL (from the ms365 MCP tool get-download-url) or base64 bytes (from ms365 download-bytes, e.g. an email attachment).
requirements: markitdown[docx,xlsx,pptx,pdf], requests
version: 0.2
"""

import base64
import os
import tempfile

import requests
from markitdown import MarkItDown

MAX_CHARS = 120_000  # keep tool output within a sensible context budget
SUPPORTED = {".docx", ".xlsx", ".xls", ".pptx", ".pdf", ".csv", ".txt", ".md", ".html"}


def _convert(data: bytes, file_name: str) -> str:
    ext = os.path.splitext(file_name)[1].lower() or ".docx"
    if ext not in SUPPORTED:
        return f"Unsupported file type '{ext}'. Supported: {', '.join(sorted(SUPPORTED))}"
    with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tmp:
        tmp.write(data)
        path = tmp.name
    try:
        text = MarkItDown().convert(path).text_content or ""
    finally:
        os.remove(path)
    if len(text) > MAX_CHARS:
        text = text[:MAX_CHARS] + f"\n\n[Truncated: showing first {MAX_CHARS} of {len(text)} characters]"
    return text


class Tools:
    def read_office_file_from_url(self, download_url: str, file_name: str) -> str:
        """
        Read a SharePoint/OneDrive Word, Excel, PowerPoint or PDF file as Markdown.
        First call the ms365 tool get-download-url with target=/drives/{drive-id}/items/{item-id}/content,
        then pass the returned URL here immediately (it expires quickly).
        :param download_url: Pre-authenticated download URL returned by get-download-url.
        :param file_name: The file's name including extension, e.g. "Q3 Review.pptx".
        :return: The document content as Markdown.
        """
        resp = requests.get(download_url, timeout=120)
        resp.raise_for_status()
        return _convert(resp.content, file_name)

    def read_office_file_from_base64(self, content_base64: str, file_name: str) -> str:
        """
        Read a Word, Excel, PowerPoint or PDF file supplied as base64 (e.g. an email attachment
        fetched with ms365 download-bytes and target=/me/messages/{id}/attachments/{att-id}/$value).
        :param content_base64: The file bytes encoded as base64.
        :param file_name: The file's name including extension, e.g. "invoice.xlsx".
        :return: The document content as Markdown.
        """
        data = base64.b64decode(content_base64)
        return _convert(data, file_name)
