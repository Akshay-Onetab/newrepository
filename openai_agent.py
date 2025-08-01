from flask import Flask, request, jsonify
from flask_cors import CORS  # Add this import
from langchain.agents import initialize_agent, AgentType
from langchain_openai import ChatOpenAI
from oneNote.oneNote_create_tools import onenote_create_tools
from oneNote.oneNote_get_tools import onenote_get_tools
from ToDo.todo_get_tools import todo_get_tools
from qdrant.qdrant_tools import qdrant_query_tool
from auth.auth import get_Token
from oneDrive.oneDrive_create_tool import onedrive_create_tools
from oneDrive.oneDrive_get_tool import onedrive_get_tools
import dotenv
import os

# Load .env
dotenv.load_dotenv()

# # Ensure access token
# if not os.getenv("GRAPH_ACCESS_TOKEN"):
#     raise ValueError("GRAPH_ACCESS_TOKEN not found in .env")

# Flask setup
app = Flask(__name__)
CORS(app, origins=["http://localhost:5173"])

# Initialize LLM & tools
llm = ChatOpenAI(model='gpt-4', temperature=0)
# tools = onenote_create_tools + onenote_get_tools + todo_get_tools + get_Token + onedrive_create_tools + onedrive_get_tools  # both are lists

qdrant_tool = qdrant_query_tool(collection_name="ONENOTE_COLLECTION")
tools = qdrant_tool + onenote_create_tools 
agent = initialize_agent(
    tools=tools,
    llm=llm,
    agent=AgentType.ZERO_SHOT_REACT_DESCRIPTION,
    verbose=True
)

@app.route("/run-agent", methods=["POST"])
def run_agent():
    data = request.get_json()
    print(data, "-------------data")
    if not data or "prompt" not in data:
        return jsonify({"error": "Missing 'prompt' in request"}), 400
    try:
        verify = data['agent']
        if verify == "Avi":
            response = agent.run(data["prompt"])
            return jsonify({"response": response})
        else:
            return "Unauthorized"
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(debug=True)
