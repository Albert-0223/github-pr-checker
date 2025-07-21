import requests

GITHUB_TOKEN = "GITHUB_TOKEN"  # Set your GitHub personal access token
headers = {"Authorization": f"token {GITHUB_TOKEN}"}

response = requests.get("https://api.github.com/user", headers=headers)

print(f"Status Code: {response.status_code}")
print("Response:", response.json())