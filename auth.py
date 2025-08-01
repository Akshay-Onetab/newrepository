from langchain.tools import Tool
import requests
import time

def get_token_with_retry(retries=4, delay=2):
    """
    Retrieve access token for Microsoft Graph API, retrying on failure.
    """
    for attempt in range(retries):
        response = requests.get('http://localhost:3002/microsoft/access-token')
        if response.status_code == 200 and response.text and "Failed" not in response.text:
            return response.text
        time.sleep(delay)
    raise Exception("Failed to fetch access token after multiple attempts.")

get_Token = [
    Tool(
        name="GetToken",
        func=lambda _: get_token_with_retry(),
        description="Retrieve access token for Microsoft Graph API when access token expires. Retries automatically."
    ),
]