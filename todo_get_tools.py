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
    "Authorization": f"Bearer {ACCESS_TOKEN}",
    "Content-Type": "application/json"
}


def get_TodoLists():
    url = "https://graph.microsoft.com/v1.0/me/todo/lists"
    response = requests.get(url, headers=HEADERS)
    if response.status_code == 200:
        todo_lists = response.json().get("value", [])
        return "\n".join([f"{todo_list['title']} (ID: {todo_list['id']})" for todo_list in todo_lists])
    return f"Failed to fetch todo lists: {response.text}"

def get_TodoTasks(todoTaskListId: str):
    url = f"https://graph.microsoft.com/v1.0/me/todo/lists/{todoTaskListId}/tasks"
    response = requests.get(url, headers=HEADERS)
    if response.status_code == 200:
        todo_tasks = response.json().get("value", [])
        return "\n".join([f"{todo_task['title']} (ID: {todo_task['id']})" for todo_task in todo_tasks])
    return f"Failed to fetch todo tasks: {response.text}"


todo_get_tools = [
    Tool(
        name="GetTodoLists",
        func=lambda _: get_TodoLists(),
        description="Retrieve all the names of todo lists for the signed-in user."
    ),
    Tool(
        name="GetTodoTasks",
        func=lambda todoTaskListId: get_TodoTasks(todoTaskListId),
        description="Retrieve all todo tasks from a specific todo list by its ID."
    )
]