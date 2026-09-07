"""
GitHub Pull Request Tool

Features:
1. Extract pull request details.
2. Update pull request description.
3. List last N pull requests.
4. Export results to JSON and XLSX.
5. Read configuration from PR description.
"""

import json
import os
import re
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
            files_changed (int): Number of files changed.
            files_added (int): Number of lines added.
            files_deleted (int): Number of lines deleted.
        """

        self.number = number
        self.title = title
        self.description = description
        self.author = author
        self.files_changed = files_changed
        self.files_added = files_added
        self.files_deleted = files_deleted

    def to_dict(self) -> dict[str, Any]:
        """
        Convert PR object into dictionary.

        Returns:
            dict[str, Any]: Pull request information.
        """

        return {
            "pr_number": self.number,
            "title": self.title,
            "description": self.description,
            "author": self.author,
            "files_changed": self.files_changed,
            "files_added": self.files_added,
            "files_deleted": self.files_deleted,
        }


class GitHubPRTool:
    """
    GitHub Pull Request management tool.
    """

    def __init__(
        self,
        token: str,
        repo_name: str,
    ) -> None:
        """
        Establish GitHub connection.

        Args:
            token (str): GitHub token.
            repo_name (str): Repository name.
        """

        auth = Auth.Token(token)

        self.github = Github(auth=auth)

        self.repo = self.github.get_repo(
            repo_name
        )

    def get_pr_details(
        self,
        pr_number: int,
    ) -> PullRequestInfo:
        """
        Retrieve pull request details.

        Args:
            pr_number (int): Pull request number.

        Returns:
            PullRequestInfo:
                Pull request details.
        """

        pr = self.repo.get_pull(
            pr_number
        )

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
        Update pull request description.

        Args:
            pr_number (int):
                Pull request number.

            new_description (str):
                Updated description.
        """

        pr = self.repo.get_pull(
            pr_number
        )

        pr.edit(
            body=new_description
        )

    def list_pull_requests(
        self,
        count: int = 10,
        state: str = "open",
    ) -> list[dict[str, Any]]:
        """
        List latest pull requests.

        Args:
            count (int):
                Number of pull requests.

            state (str):
                open / closed / all.

        Returns:
            list[dict[str, Any]]
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
        Export data into JSON.

        Args:
            data (list):
                Pull request data.

            output_file (str):
                Output file name.
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
        Export data into Excel.

        Args:
            data (list):
                Pull request data.

            output_file (str):
                Output file name.
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

        workbook.save(
            output_file
        )

    def parse_pr_description(
        self,
        pr_description: str,
        current_repo: str,
    ) -> tuple[int, str]:
        """
        Read PR description and
        extract configuration values.

        Args:
            pr_description (str):
                PR description.

            current_repo (str):
                Current repository.

        Returns:
            tuple[int, str]
        """

        number_of_prs = 10

        repo_name = current_repo

        count_match = re.search(
            r"number_of_prs=(\d+)",
            pr_description,
            re.IGNORECASE,
        )

        if count_match:
            number_of_prs = int(
                count_match.group(1)
            )

        repo_match = re.search(
            r"repo_url=https://github\.com/([^\s]+)",
            pr_description,
            re.IGNORECASE,
        )

        if repo_match:
            repo_name = repo_match.group(1)

        return (
            number_of_prs,
            repo_name,
        )

    def get_pr_configuration(
        self,
        pr_number: int,
    ) -> tuple[int, str]:
        """
        Get configuration values
        from PR description.

        Args:
            pr_number (int):
                Pull request number.

        Returns:
            tuple[int, str]
        """

        pr = self.repo.get_pull(
            pr_number
        )

        description = pr.body or ""

        return self.parse_pr_description(
            description,
            self.repo.full_name,
        )

    def write_pr_details_excel(
        self,
        pr_details: PullRequestInfo,
        output_file: str,
    ) -> None:
        """
        Export detailed pull request data
        to Excel.
        """

        workbook = Workbook()

        worksheet = workbook.active

        worksheet.title = "PR Details"

        worksheet.append(
            [
                "PR Number",
                "Title",
                "Description",
                "Author",
                "Files Changed",
                "Files Added",
                "Files Deleted",
            ]
        )

        worksheet.append(
            [
                pr_details.number,
                pr_details.title,
                pr_details.description,
                pr_details.author,
                pr_details.files_changed,
                pr_details.files_added,
                pr_details.files_deleted,
            ]
        )

        workbook.save(output_file)


def main() -> None:
    """
    Application entry point.
    """

    token = os.getenv(
        "GITHUB_TOKEN"
    )

    if not token:
        print(
            "GITHUB_TOKEN not found."
        )
        return

    current_repo = (
        "SpoorthiR-2511/GitHubPRTool"
    )

    tool = GitHubPRTool(
        token,
        current_repo,
    )

    pr_number = 1

    # Requirement 1
    pr_details = tool.get_pr_details(
        pr_number
    )

    print("\nPR DETAILS")
    print(
        f"Title: {pr_details.title}"
    )
    print(
        f"Author: {pr_details.author}"
    )
    print(
        f"Description: {pr_details.description}"
    )
    print(
        f"Files Changed: {pr_details.files_changed}"
    )
    print(
        f"Files Added: {pr_details.files_added}"
    )
    print(
        f"Files Deleted: {pr_details.files_deleted}"
    )

    tool.write_json(
        [pr_details.to_dict()],
        "pr_details.json",
    )

    # Requirement 5
    count, repo_name = (
        tool.get_pr_configuration(
            pr_number
        )
    )

    print(
        f"\nRepository: {repo_name}"
    )

    print(
        f"Number Of PRs: {count}"
    )

    tool.repo = (
        tool.github.get_repo(
            repo_name
        )
    )

    prs = tool.list_pull_requests(
        count=count,
        state="open",
    )

    tool.write_json(
        prs,
        "prs.json",
    )

    tool.write_excel(
        prs,
        "prs.xlsx",
    )
    tool.write_pr_details_excel(
        pr_details,
        "pr_details.xlsx",
    )

    print(
        "\nReports generated successfully."
    )


if __name__ == "__main__":
    main()
