from langchain.tools import Tool
import requests
import os
from bs4 import BeautifulSoup

# Load token from env or secret manager
try:
    ACCESS_TOKEN = requests.get('http://localhost:3000/microsoft/access-token').text
except:
    ACCESS_TOKEN = os.getenv("GRAPH_ACCESS_TOKEN")
HEADERS = {
    "Authorization": f"Bearer EwC4BMl6BAAUBKgm8k1UswUNwklmy2v7U/S+1fEAAfpZ5kGgJDxxSsQ4weI48vpdHAyrmoqA5FaqVNSsOxBEmBuGZDQTBnTNnzVBkvf+KCnzSnKoG17HKtr77OPzgZY4nxzMiCKXJqdPc2wh4/ta7zDItJ2+Xno6t5VoNEidaT3bIWwafmyzfonVh7N290PllyIpRxo7fQk57nBcKv7VSvcWEUCTpMtLhv4kVMwglZmYvygEWKbspr3zal6Gv6fje/moEfW6ODf1sJPTwCIVuzITVQihrcougUef7ZXrDlPmb9poOExuA9U5bxxqCXjW94KqOoAqnL/UD8KsW+JCd0IWstIcNSGc+AODDQrWmp5my6s2tEakd1gQK7aSeVIQZgAAEK2VjzHiixZXIwbijH0+5UeAA2nbM9dmTph4+0ZdE2jPWh+sQX1daTJ7FO4j+pq9cEv7TAbBBEpYY2fISZ4kKyKLwC93r/nCODuvDVmzfgQxPDqbiAGjaVZEToNcLuQO09wOR0BWVlTs7fm4CPDRAbZyyaeeF1BSo2cArgyFMPJ0PnCzmWR1HujnbT46BwqZVWGONulGzsPpsvoD3ptu3Raz+xQuzDWGGK/fNa2136iVsK8xw63JeQVtMdyhuOYWS8fOj3ES9+qgVBifhOXcofBwSwiZ3ntxeLzMWiLcuyMYlyk1kFl8AAJ4axx63gifAHMD+pWBDYZyQmpVUWHTpajkUvyj1QteMSQwu4VkQ6kJU9yFM38kbLZaMK/IiaXuwDhG5VXJi9HVnoYhOt2bqrNBFg28EPgFdl8EX/rrttXRDTThRFd1cSUlnCrv5Cv/J4t0WlFb9w/RqQawACX/qlQphitKtKIlGe5/TWBbkytwVP2SZh8QSqVG/AgqYo/MzKd1jFVfMbJ9bIh0kLPgIuYQlB7dJUeqCb3S5HjhTN5cdQ+X0ZtCu7bg+sREi+UayHY1nLzXMnYcorMsW/W1ZL3Z2iATxTDn1aFtj/7pvXCz/rYSEzqvlerJAL4NfX9OZ4o6L/HWUKRdlAsj2EmETfuR+5Wnf3draOE02y6AYVx3RqsvvP/PKJMWGKaYThR05ITgGRhnUmhri5kIIj3+Z36GdEu1nJg094pX3pUhg0Qow5XIEGyv7A2zd2mkB0V36X2w6dPsBP8UgszyCNL4Yo+L9Zz6w27gE9FkPCXvwGzmHLms/nw0M6RnJo+X00vXlGrOsKclazTao9rdH+tLEk58CwJfD3h8ndxE8qKlBMesSYhhwgOph+bIVGZL73NwuxERDvwSuWVeP6OvlswXnKA8m8oWdeMDTuW/OGinGyucrZ5BOjJm/3BR5YdLrYUbz+FM3ZhqMHY47giRXnMnjsf/YAkTLhh8CA6uOcDAGQgwNfp6xFalP8eg+DahC5MBxSIJM09Hj7rpKLYirvaVsBGAa5n8LVpwzd6264EXJwRXr4oYAkHb6PL/mQTKLC3g1Gi/b7b+5O7Dh+eXF4EebFJjKw7wsoYAAWjaFd7ic7gSrPZe8ejJ5zB54fBpRZz5rqT7+kEpEyYrJliDur1HwAZvU7vRY2YoAOb+OW/X2aOFcnZ+hOKd0bihuL2r1Q6D3cAEzgM=",
    "Content-Type": "application/json"
}

# -------------------------
# Get Notebooks
# -------------------------
def get_notebooks():
    url = "https://graph.microsoft.com/v1.0/me/onenote/notebooks"
    response = requests.get(url, headers=HEADERS)
    if response.status_code == 200:
        notebooks = response.json().get("value", [])
        return "\n".join([f"{nb['displayName']} (ID: {nb['id']})" for nb in notebooks])
    return f"Failed to fetch notebooks: {response.text}"

# -------------------------
# Get Sections from a Specific Notebook
# -------------------------
def get_sections(notebook_id: str):
    url = f"https://graph.microsoft.com/v1.0/me/onenote/notebooks/{notebook_id}/sections"
    response = requests.get(url, headers=HEADERS)
    if response.status_code == 200:
        sections = response.json().get("value", [])
        return "\n".join([f"{sec['displayName']} (ID: {sec['id']})" for sec in sections])
    return f"Failed to fetch sections: {response.text}"

# -------------------------
# Get all Sections
# -------------------------
def get_all_sections():
    url = "https://graph.microsoft.com/v1.0/me/onenote/sections"
    response = requests.get(url, headers=HEADERS)
    if response.status_code == 200:
        sections = response.json().get("value", [])
        return "\n".join([f"{sec['displayName']} (ID: {sec['id']})" for sec in sections])
    return f"Failed to fetch sections: {response.text}"

# -------------------------
# Get Pages from a Specific Section
# -------------------------
def get_pages(section_id: str):
    url = f"https://graph.microsoft.com/v1.0/me/onenote/sections/{section_id}/pages"
    response = requests.get(url, headers=HEADERS)
    if response.status_code == 200:
        pages = response.json().get("value", [])
        return "\n".join([f"{page['title']} (ID: {page['id']})" for page in pages])
    return f"Failed to fetch pages: {response.text}"

# -------------------------
# Get all Pages
# -------------------------
def get_all_pages():
    url = "https://graph.microsoft.com/v1.0/me/onenote/pages"
    response = requests.get(url, headers=HEADERS)
    if response.status_code == 200:
        pages = response.json().get("value", [])
        return "\n".join([f"{page['title']} (ID: {page['id']})" for page in pages])
    return f"Failed to fetch pages: {response.text}"

# ------------------
# Get Pages Content
# ------------------
def get_page_content(page_id: str):
    url = f"https://graph.microsoft.com/v1.0/me/onenote/pages/{page_id}/content"
    response = requests.get(url, headers=HEADERS)
    if response.status_code == 200:
        return response.text
    return f"Failed to fetch page content: {response.text}"

onenote_get_tools = [
    Tool(
        name="GetNotebooks",
        func=lambda _: get_notebooks(),
        description="Retrieve all notebooks for the signed-in user."
    ),
    Tool(
        name="GetSections",
        func=lambda notebook_id: get_sections(notebook_id),
        description="Retrieve all sections in a specific notebook. Input should be the notebook ID."
    ),
    Tool(
        name="GetAllSections",
        func=lambda _: get_all_sections(),
        description="Retrieve all sections across all notebooks."
    ),
    Tool(
        name="GetPages",
        func=lambda section_id: get_pages(section_id),
        description="Retrieve all pages in a specific section. Input should be the section ID."
    ),
    Tool(
        name="GetAllPages",
        func=lambda _: get_all_pages(),
        description="Retrieve all pages across all sections."
    ),
    Tool(
        name="GetPageContent",
        func=lambda page_id: get_page_content(page_id),
        description="Get content of a OneNote page. Input should be the page ID."
        # description="Get the exact raw HTML content of a OneNote page. Input should be the page ID. Return the HTML as-is, without summarizing or interpreting."
    )
]
