"""
GitHub Pull Request Tool
"""

import argparse
import json
import os
from typing import Any

from github import Auth
from github import Github
from openpyxl import Workbook


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
        Initialize GitHub connection.
        """

        auth = Auth.Token(token)

        self.github = Github(
            auth=auth
        )

        self.repo = (
            self.github.get_repo(
                repo_name
            )
        )

    def get_pr_details(
        self,
        pr_number: int,
    ) -> dict[str, Any]:
        """
        Get pull request details.
        """

        pr = self.repo.get_pull(
            pr_number
        )


        return {
            "pr_number": pr.number,
            "title": pr.title,
            "description": pr.body or "",
            "author": (
                pr.user.login
                if pr.user
                else "Unknown"
            ),
            "files_changed": pr.changed_files,
        }

    def update_pr_description(
        self,
        pr_number: int,
        new_description: str,
    ) -> None:
        """
        Update PR description.
        """

        pr = self.repo.get_pull(
            pr_number
        )

        pr.edit(
            body=new_description
        )

    def list_pull_requests(
        self,
        count: int,
        state: str,
    ) -> list[dict[str, Any]]:
        """
        List pull requests.
        """

        results = []

        if state == "merged":

            pulls = self.repo.get_pulls(
                state="closed",
            )

        else:

            pulls = self.repo.get_pulls(
                state=state,
            )

        counter = 0

        for pr in pulls:

            if state == "merged" and not pr.merged:
                continue

            if counter >= count:
                break

            counter += 1

            results.append(
                {
                    "pr_number": pr.number,
                    "title": pr.title,
                    "description": (
                        pr.body or ""
                    ),
                    "author": (
                        pr.user.login
                        if pr.user
                        else "Unknown"
                    ),
                    "files_changed": (
                        pr.changed_files
                    ),
                }
            )

        return results

def write_json(
    data: list[dict[str, Any]],
    output_file: str,
) -> None:
    """
    Write JSON report.
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
    data: list[dict[str, Any]],
    output_file: str,
) -> None:
    """
    Write Excel report.
    """

    workbook = Workbook()

    worksheet = (
        workbook.active
    )

    worksheet.title = (
        "Pull Requests"
    )vv

    worksheet.append(
        [
            "PR Number",
            "Title",
            "Description",
            "Author",
            "Files Changed",
        
        ]
    )

    for row in data:

        worksheet.append(
            [
                row["pr_number"],
                row["title"],
                row["description"],
                row["author"],
                row["files_changed"],
              
            ]
        )

    workbook.save(
        output_file
    )


def parse_arguments() -> (
    argparse.Namespace
):
    """
    Parse CLI arguments.
    """

    parser = (
        argparse.ArgumentParser()
    )

    parser.add_argument(
        "--repo",
        required=True,
    )

    parser.add_argument(
        "--pr_number",
        type=int,
    )

    parser.add_argument(
        "--get_pr",
        action="store_true",
    )

    parser.add_argument(
        "--update_pr_description",
        action="store_true",
    )

    parser.add_argument(
        "--new_desc",
    )

    parser.add_argument(
        "--list_pr",
        action="store_true",
    )

    parser.add_argument(
        "--numbers",
        type=int,
        default=10,
    )

    parser.add_argument(
        "--status",
        default="open",
    )

    return (
        parser.parse_args()
    )


def main() -> None:
    """
    Application entry point.
    """

    try:

        token = os.getenv(
            "GITHUB_TOKEN"
        )

        if not token:
            raise ValueError(
                "GITHUB_TOKEN "
                "environment "
                "variable not found."
            )

        args = parse_arguments()

        tool = GitHubPRTool(
            token,
            args.repo,
        )

        if args.get_pr:

            if args.pr_number is None:
                raise ValueError(
                    "--pr_number is required with "
                    "--get_pr"
                )

            data = [
                tool.get_pr_details(
                    args.pr_number
                )
            ]

        elif args.update_pr_description:

            if args.pr_number is None:
                raise ValueError(
                    "--pr_number is required with "
                    "--update_pr_description"
                )

            if not args.new_desc:
                raise ValueError(
                    "--new_desc is required with "
                    "--update_pr_description"
                )

            tool.update_pr_description(
                args.pr_number,
                args.new_desc,
            )

            print(
                "PR description "
                "updated successfully."
            )

            return

        elif args.list_pr:

            if args.numbers <= 0:
                raise ValueError(
                    "--numbers must be greater than 0"
                )

            data = (
                tool.list_pull_requests(
                    count=args.numbers,
                    state=args.status,
                )
            )

        else:

            raise ValueError(
                "Please select one operation: "
                "--get_pr, "
                "--update_pr_description "
                "or --list_pr"
            )

        write_json(
            data,
            "report.json",
        )

        write_excel(
            data,
            "report.xlsx",
        )

        print(
            "Reports generated "
            "successfully."
        )

    except Exception as error:

        print(
            f"Error: {error}"
        )


if __name__ == "__main__":
    main()