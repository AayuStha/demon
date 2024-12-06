import requests
from datetime import datetime, timedelta
from dotenv import load_dotenv
import json
import pytz
import re
import os
load_dotenv()
API_KEY = os.getenv("API_KEY")


# Your GitHub Personal Access Token (replace with your actual token)
GITHUB_ACCESS_TOKEN = "API_KEY"

# Function to retrieve first and last commit dates
def get_first_and_last_commit(repo_owner, repo_name):
    url = f'https://api.github.com/repos/{repo_owner}/{repo_name}/commits'
    headers = {
        "Authorization": f"token {GITHUB_ACCESS_TOKEN}"
    }
    params = {'per_page': 100, 'page': 1}
    commits = []

    while True:
        response = requests.get(url, headers=headers, params=params)
        if response.status_code == 200:
            new_commits = response.json()
            if not new_commits:
                break
            commits.extend(new_commits)
            params['page'] += 1
        elif response.status_code == 403:  # Rate limit exceeded or unauthorized
            return None, None, "Rate limit exceeded or unauthorized"
        else:
            return None, None, f"Error: {response.status_code}"

    if not commits:
        return None, None, "No commits found"

    first_commit = commits[-1]['commit']['author']['date']
    last_commit = commits[0]['commit']['author']['date']
    return len(commits), first_commit, last_commit

# Convert UTC timestamp to Kathmandu timezone
def convert_utc_to_kathmandu(utc_timestamp):
    utc_datetime = datetime.strptime(utc_timestamp, '%Y-%m-%dT%H:%M:%SZ')
    utc_datetime = utc_datetime.replace(tzinfo=pytz.utc)
    kathmandu_timezone = pytz.timezone('Asia/Kathmandu')
    return utc_datetime.astimezone(kathmandu_timezone)

# Extract username and repository name from GitHub URL
def extract_repo_user(url):
    pattern = r'https://github.com/([^/]+)/([^/?]+?)(?:\.git)?$'
    match = re.match(pattern, url)
    if match:
        return match.group(1), match.group(2)
    return None, None

if __name__ == "__main__":
    # Define hackathon times
    kathmandu_timezone = pytz.timezone('Asia/Kathmandu')
    hackathon_start = kathmandu_timezone.localize(datetime(2024, 11, 28, 9, 0, 0))
    hackathon_end = hackathon_start + timedelta(hours=48)

    github_urls = [
        "https://github.com/KCNikhil/NIKHIL-KC.git",
        "https://github.com/SudeepKarki10/mappingproject"
    ]
    outputs = []

    try:
        for github_url in github_urls:
            username, repo = extract_repo_user(github_url)
            if not username or not repo:
                outputs.append({
                    "repo": None,
                    "username": None,
                    "status": "Invalid URL"
                })
                continue

            try:
                total_commits, first_commit, last_commit = get_first_and_last_commit(username, repo)
                if not total_commits:
                    outputs.append({
                        "repo": repo,
                        "username": username,
                        "status": first_commit  # Error message
                    })
                    continue

                # Convert commit times to Kathmandu timezone
                first_commit_kathmandu = convert_utc_to_kathmandu(first_commit)
                last_commit_kathmandu = convert_utc_to_kathmandu(last_commit)

                # Check if commits exist before hackathon
                has_committed_before_hackathon = first_commit_kathmandu < hackathon_start

                # Count commits during hackathon
                commits_during_hackathon = 0
                if first_commit_kathmandu <= hackathon_end:
                    commits_during_hackathon = sum(
                        hackathon_start <= convert_utc_to_kathmandu(commit['commit']['author']['date']) <= hackathon_end
                        for commit in requests.get(
                            f"https://api.github.com/repos/{username}/{repo}/commits",
                            headers={"Authorization": f"token {GITHUB_ACCESS_TOKEN}"},
                            params={"per_page": 100, "page": 1}
                        ).json()
                    )

                # Append output
                outputs.append({
                    "repo": repo,
                    "username": username,
                    "commits": total_commits,
                    "first_commit_kathmandu": first_commit_kathmandu.isoformat(),
                    "last_commit_kathmandu": last_commit_kathmandu.isoformat(),
                    "has_committed_before_hackathon": has_committed_before_hackathon,
                    "commits_during_hackathon": commits_during_hackathon,
                    "hackathon_start": hackathon_start.isoformat(),
                    "hackathon_end": hackathon_end.isoformat(),
                    "status": 200
                })
            except Exception as e:
                outputs.append({
                    "repo": repo,
                    "username": username,
                    "status": f"Error: {str(e)}"
                })
    finally:
        with open("outputs.json", "w") as f:
            json.dump(outputs, f, ensure_ascii=False, indent=6)
