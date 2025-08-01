from qdrant_client import QdrantClient
from qdrant_client.models import Distance, VectorParams, PointStruct
from qdrant.embeddings import generate_embeddings
import uuid


client = QdrantClient(url="http://localhost:6333")

def setup_qdrant(collection_name, vector_size):
    """
    Create a Qdrant collection for storing embeddings.
    """
    client.delete_collection(collection_name)  # Delete the existing collection
    client.create_collection(
        collection_name=collection_name,
        vectors_config=VectorParams(size=vector_size, distance=Distance.COSINE),
    )

def store_embeddings(collection_name, texts, metadata_list):
    """
    Store embeddings in Qdrant without overwriting existing data.
    """
    embeddings = generate_embeddings(texts)
    print(f"Generated {len(embeddings)} embeddings for {len(texts)} texts.")
    
    # Generate unique IDs for each point
    points = [
        PointStruct(id=str(uuid.uuid4()), vector=embeddings[i], payload=metadata_list[i])
        for i in range(len(texts))
    ]
    
    # Upsert points into Qdrant
    client.upsert(collection_name=collection_name, points=points)

def query_embeddings(collection_name, query_vector, top_k=5):
    """
    Query Qdrant for the closest embeddings.
    """
    results = client.search(
        collection_name=collection_name,
        query_vector=query_vector,
        limit=top_k,
    )
    return results

# print("Qdrant client initialized successfully.")

# client.create_collection(
#     collection_name="test_collection",
#     vectors_config=VectorParams(size=4, distance=Distance.DOT),
# )

# print("Collection 'test_collection' created successfully.")

# operation_info = client.upsert(
#     collection_name="test_collection",
#     wait=True,
#     points=[
#         PointStruct(id=1, vector=[0.05, 0.61, 0.76, 0.74], payload={"city": "Berlin"}),
#         PointStruct(id=2, vector=[0.19, 0.81, 0.75, 0.11], payload={"city": "London"}),
#         PointStruct(id=3, vector=[0.36, 0.55, 0.47, 0.94], payload={"city": "Moscow"}),
#         PointStruct(id=4, vector=[0.18, 0.01, 0.85, 0.80], payload={"city": "New York"}),
#         PointStruct(id=5, vector=[0.24, 0.18, 0.22, 0.44], payload={"city": "Beijing"}),
#         PointStruct(id=6, vector=[0.35, 0.08, 0.11, 0.44], payload={"city": "Mumbai"}),
#     ],
# )

# print(operation_info)