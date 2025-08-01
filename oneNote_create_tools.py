from langchain.tools import Tool
import requests
import os
import time

# Load token from env or secret manager
try:
    ACCESS_TOKEN = requests.get('http://localhost:3000/microsoft/access-token').text
except:
    ACCESS_TOKEN = os.getenv("GRAPH_ACCESS_TOKEN")
HEADERS = {
    "Authorization": f"Bearer EwBIBMl6BAAUBKgm8k1UswUNwklmy2v7U/S+1fEAAZFvjTh2XcO8lWON7kkzB5p7/8ZIL+En7uNjvM7IhSb8w4AyeegkQrCc1hc0s4B6/7vZUFwaCXzDg4jotfMDfauMEEDilTak8UzGvdCUeQIysBR4E46HT2DV3vUGUtLo4zDhJwiCiwM8FRbymgs2acaw2wXgURd3eBvl1GoT14oZePjt7RyEpww0TXPzcy6sg1yQWdMvo0glyJPXFzRSx0psQynAnoJw/SUcM3Whd/+8dBvtVDDbk0po+S5mVvrrrxUU7js56VtcpDExl5Ce80czbGoOvpNBBKTAXWHViLzqCe1TyUT8nNcggBh8e5RmhTwkS4VCUD/s+XK2gSe40CQQZgAAEGg7LYsD+JAigtEr1cp08o0QA5nFR3hEQ5Iv9uPt3+zUjPPNbPN6rFCTdxYi/m7ClzT0/4liWIoiFSWe3x3nX+IXH4qOrITlgoQGQjmHxYgHSFTEaFWeUPiZYaN+E2BngigXQp+g7x3sATObcExfRfU2NJrEDQEVLn7N/X2tlwrDhqsEQpYl4Q/8jMWZRB6opXjnOGQedCvJRyM3ZMDbZsNWPQyUbpQ/cymWhiIMT8TgzxbNkGLQVupLYf3qaJIzitbX4igEkk/t5vhCydgjtVsbulSb2kiH2lKxyoknIU+uJu0KQhbOQk5JvEgFmOcqxTUYpMoh4+jX6/ERx3McGG/kqJSJTkGUgaq0dyRlfFW1wNqiRuEUbvU4JbyU+EXSqDTqPWxwpO+T9Dyds8NSN4LEEPJym4OpgoXN2T36KJS+rAIR7tmYm4/gVlzQGBvu8/JSCMstKKXItthTfgEhFDrh+5X1x7ptAkL1w5BGCwbPJgF5fb9tnjkjEqn0Es6IlRfHMAOYXEF8SFQ5XMH7gtWz5vra4hexAF7vzj/39rUhl7lCTuQG0EVsvfoDWeCtlVuMHI/6ScDDNYCQTRWGyKxKE0wE+GD860Fwt6OIbt1qoynjkbLYzAr29jajbJcBJb2OuykZzvVub0G1/UNZk2WAi2MJDYNZWrmHxci8EVbt9EaqfDpo3diGIUGypTfoK4/PZXVFLatp2Mr1G0gvUh4GH/nuoINV+AjT1qMPzYTG47yVDK5LV3itxsZB0I45jCuF5Murrk2fGfINsJYFZFqc+gr7k2XzpMy9ZMgjbyJJsdKdh/dxCg7s8rrhl5B2FZ90MKf4JVdzf4eTEZpxuDDU4Exqiud3YzMbsZm1VBnexUz44gEhQJ+3gvD+OJ5j36PvxAK0sVB83mMh9IQGLRaqzYlyftLMVt3qIcYBComP7Roogpk5EjDyx0KS0H2ObZz/lF5p/oTfFyisgf1ZbPD8mX6uyAx5GfxTGUVNmOdjJ33opCCgUm+jE/DjvWckYFSTcmcy/nnwk0pts15LHLswqiyoImsGgodBCuA0V9ZN4p5LAw==",
    "Content-Type": "application/json"
}

# -------------------------
# Create a new OneNote Notebook
# -------------------------
def create_notebook(name: str):
    url = "https://graph.microsoft.com/v1.0/me/onenote/notebooks"
    payload = {"displayName": name}
    response = requests.post(url, headers=HEADERS, json=payload)
    if response.status_code == 201:
        return f"Notebook '{name}' created successfully."
    return f"Failed to create notebook: {response.text}"

# -------------------------
# Create a new Section inside a Notebook
# -------------------------
def create_section(notebook_id: str, section_name: str, retries=3, delay=2):
    url = f"https://graph.microsoft.com/v1.0/me/onenote/notebooks/{notebook_id}/sections"
    payload = {"displayName": section_name}
    for attempt in range(retries):
        response = requests.post(url, headers=HEADERS, json=payload)
        if response.status_code == 201:
            return f"Section '{section_name}' created in notebook {notebook_id}."
        # If transient error, retry
        try:
            error_json = response.json()
            if (
                response.status_code == 500 or
                error_json.get("error", {}).get("code") == "20280"
            ):
                if attempt < retries - 1:
                    time.sleep(delay)
                    continue
        except Exception:
            pass
        return f"Failed to create section: {response.text}"
    return f"Failed to create section after {retries} attempts: {response.text}"

# -------------------------
# Create a new Page inside a Section
# -------------------------
def create_page(section_id: str, title: str, body: str):
    url = f"https://graph.microsoft.com/v1.0/me/onenote/sections/{section_id}/pages"
    html_body = f"""
        <html>
            <head><title>{title}</title></head>
            <body>
                <p>{body}</p>
            </body>
        </html>
    """
    headers = HEADERS.copy()
    headers["Content-Type"] = "application/xhtml+xml"
    response = requests.post(url, headers=headers, data=html_body.encode("utf-8"))
    if response.status_code == 201:
        return f"Page '{title}' created in section {section_id}."
    return f"Failed to create page: {response.text}"

# -------------------------
# LangChain Tools
# -------------------------
onenote_create_tools = [
    Tool(
        name="CreateNotebook",
        func=lambda name: create_notebook(name),
        description="Create a new OneNote notebook. Input should be the notebook name."
    ),
    Tool(
        name="CreateSection",
        func=lambda args: create_section(*args.split("|")),
        description="Create a section in a notebook. Input should be 'notebook_id|section_name'."
    ),
    Tool(
        name="CreatePage",
        func=lambda args: create_page(*args.split("|")),
        description="Create a page in a section. Input should be 'section_id|page_title|page_body'. If page_title or page_body is missing, generate them yourself. If the prompt includes 'summarize' or 'detail', create the content accordingly."
    )
]
