from langchain.tools import Tool
from qdrant.main import query_embeddings
import os
from qdrant.embeddings import generate_embeddings
import requests

try:
    ACCESS_TOKEN = requests.get('http://localhost:3000/microsoft/access-token').text
except:
    ACCESS_TOKEN = os.getenv("GRAPH_ACCESS_TOKEN")
HEADERS = {
    "Authorization": f"Bearer {ACCESS_TOKEN}",
    "Content-Type": "application/json"
}

def qdrant_query_tool(collection_name):
    """
    LangChain tool to query Qdrant.
    """
    def query_tool(query_text):
        query_vector = generate_embeddings([query_text])[0]
        # print(f"Querying Qdrant with vector: {query_vector}")
        results = query_embeddings(collection_name, query_vector)
        # print(f"Found {len(results)} results in Qdrant for query: {query_text}")
        # print("Results:", results[0].payload.get("paragraph", "No text found in payload"))
        # return "\n".join([str(result[0].payload) for result in results])
        return results[0].payload.get("paragraph", "No text found in payload")
    
    def get_pages_id(query_text):
        query_vector = generate_embeddings([query_text])[0]
        # print(f"Querying Qdrant with vector: {query_vector}")
        results = query_embeddings(collection_name, query_vector)
        # print(f"Found {len(results)} results in Qdrant for query: {query_text}")
        # print("Results:", results[0].payload.get("paragraph", "No text found in payload"))
        # return "\n".join([str(result[0].payload) for result in results])
        return results[0].payload.get("pages", "No text found in payload")
    
    def get_page_content(page_id: str):
        url = f"https://graph.microsoft.com/v1.0/me/onenote/pages/{page_id}/content"
        response = requests.get(url, headers=HEADERS)
        if response.status_code == 200:
            return response.text
        return f"Failed to fetch page content: {response.text}"

    return[ 
        Tool(
        name="search and retrieve notebooks,sections,pages from Qdrant",
        func=query_tool,
        description="This tool searches Qdrant for the most relevant embeddings based on the input query. It prioritizes retrieving results from the vector store to provide the most contextually relevant paragraph or data for the query. Use this tool for any search-related tasks.",
        ),
        Tool(
        name="search and retrieve pages from Qdrant",
        func=get_pages_id,
        description="This tool retrieves the IDs of pages from Qdrant based on the input query. It is useful for finding specific pages related to the query.",
        ),
        Tool(
        name="GetPageContent",
        func=lambda page_id: get_page_content(page_id),
        description="Get content of a OneNote page. Input should be the page ID."
        # description="Get the exact raw HTML content of a OneNote page. Input should be the page ID. Return the HTML as-is, without summarizing or interpreting."
         )
    ]
