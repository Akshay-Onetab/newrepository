# from langchain_openai import AzureOpenAIEmbeddings
from langchain_openai import OpenAIEmbeddings

import os

# def generate_embeddings(texts):
#     """
#     Generate embeddings for a list of texts using Azure OpenAI.
#     """
#     embeddings = AzureOpenAIEmbeddings(
#     azure_endpoint=f"https://{os.getenv('AZURE_OPENAI_API_INSTANCE_NAME')}.openai.azure.com/",
#     azure_deployment="text-embedding-ada-002",
#     openai_api_version=os.getenv("AZURE_OPENAI_API_VERSION_EMBED"),
#     api_key=os.getenv("AZURE_OPENAI_EMBED_API_KEY"),
#     model="text-embedding-ada-002",
#     )
#     return embeddings.embed_documents(texts)


def generate_embeddings(texts):
    """
    Generate embeddings for a list of texts using Azure OpenAI.
    """
    embeddings = OpenAIEmbeddings(
        # deployment=os.getenv("AZURE_OPENAI_API_DEPLOYMENT_NAME"),
        model="text-embedding-ada-002",  # Use an embedding model
        openai_api_key=os.getenv("OPENAI_API_KEY"),
        # openai_api_base=f"https://{os.getenv('AZURE_OPENAI_API_INSTANCE_NAME')}.openai.azure.com/",
        # openai_api_version=os.getenv("AZURE_OPENAI_API_VERSION"),
    )
    return embeddings.embed_documents(texts)