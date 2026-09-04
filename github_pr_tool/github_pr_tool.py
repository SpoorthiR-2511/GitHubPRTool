"""
GitHub Pull Request Tool

Features:
1. Extract pull request details.
2. Update pull request description.
3. List last N pull requests.
4. Export results to JSON and XLSX.
"""

import json
import os
from typing import Any

from github import Auth, Github
from openpyxl import Workbook

class PullRequestInfo:
    """
    Represents a GitHub Pull Request.
    """

    def __init__(
        self,
        number: int,
        title: str,
        description: str,
        author: str,
        files_changed: int,
        files_added: int,
        files_deleted: int,
    ) -> None:
        """
        Initialize pull request information.

        Args:
            number (int): PR number.
            title (str): PR title.
            description (str): PR description.
            author (str): PR author.
            files_changed (int): Files changed.
            files_added (int): Lines added.
            files_deleted (int): Lines deleted.
        """

        self.number = number
        self.title = title
        self.description = description
        self.author = author
        self.files_changed = files_changed
        self.files_added = files_added
        self.files_deleted = files_deleted




class GitHubPRTool:
    """
    GitHub Pull Request operations.
    """

    def __init__(
        self,
        token: str,
        repo_name: str,
    ) -> None:
        """
        Connect to GitHub.

        Args:
            token (str): GitHub token.
            repo_name (str): owner/repo.
        """

       
        auth = Auth.Token(token)
        self.github = Github(auth=auth)
        

    def get_pr_details(
        self,
        pr_number: int,
    ) -> PullRequestInfo:
        """
        Get details of a PR.

        Args:
            pr_number (int): Pull request number.

        Returns:
            PullRequestInfo
        """

        pr = self.repo.get_pull(pr_number)

        additions = 0
        deletions = 0

        for file in pr.get_files():
            additions += file.additions
            deletions += file.deletions

        return PullRequestInfo(
            number=pr.number,
            title=pr.title,
            description=pr.body or "",
            author=pr.user.login,
            files_changed=pr.changed_files,
            files_added=additions,
            files_deleted=deletions,
        )

    def update_pr_description(
        self,
        pr_number: int,
        new_description: str,
    ) -> None:
        """
        Update PR description.

        Args:
            pr_number (int): PR number.
            new_description (str): New description.
        """

        pr = self.repo.get_pull(pr_number)

        pr.edit(
            body=new_description,
        )

    def list_pull_requests(
        self,
        count: int = 10,
        state: str = "open",
    ) -> list[dict[str, Any]]:
        """
        List latest PRs.

        Args:
            count (int): Number of PRs.
            state (str): open/closed/all.

        Returns:
            list
        """

        results = []

        pulls = self.repo.get_pulls(
            state=state
        )

        for index, pr in enumerate(pulls):
            if index >= count:
                break

            results.append(
            {
                "pr_number": pr.number,
                "title": pr.title,
                "author": pr.user.login,
            }
        )

        return results

    def write_json(
        self,
        data: list[dict[str, Any]],
        output_file: str,
    ) -> None:
        """
        Export PR data to JSON.

        Args:
            data (list): PR data.
            output_file (str): JSON filename.
        """

        with open(
            output_file,
            "w",
            encoding="utf-8",
        ) as file:
            json.dump(
                data,
                file,
                indent=4,
            )

    def write_excel(
        self,
        data: list[dict[str, Any]],
        output_file: str,
    ) -> None:
        """
        Export PR data to XLSX.

        Args:
            data (list): PR data.
            output_file (str): Excel filename.
        """

        workbook = Workbook()

        worksheet = workbook.active

        worksheet.title = "Pull Requests"

        worksheet.append(
            [
                "PR Number",
                "Title",
                "Author",
            ]
        )

        for row in data:
            worksheet.append(
                [
                    row["pr_number"],
                    row["title"],
                    row["author"],
                ]
            )

        workbook.save(output_file)


def main() -> None:
    """
    Application entry point.
    """

    token = os.getenv("GITHUB_TOKEN")

    if not token:
        print("GITHUB_TOKEN not found.")
        return

    repo_name = "SpoorthiR-2511/GitHubPRTool"

    tool = GitHubPRTool(
        token,
        repo_name,
    )

    prs = tool.list_pull_requests(
        count=10,
        state="all",
    )

    if not prs:
        print("No pull requests found.")
        return

    tool.write_json(
        prs,
        "prs.json",
    )

    tool.write_excel(
        prs,
        "prs.xlsx",
    )

    print("Reports generated successfully.")


if __name__ == "__main__":
    main()