import argparse
import logging
import re
from dataclasses import dataclass
from typing import Any

import xlsxwriter
import yaml
from docutils.core import publish_doctree
from docutils.nodes import (
    GenericNodeVisitor,
    Node,
    Text,
    document,
    inline,
    literal,
    literal_block,
    paragraph,
    strong,
    system_message,
    title_reference,
)

logger = logging.getLogger()


@dataclass(frozen=True)
class Qcm:
    title: str
    correct: frozenset[int]
    choices: tuple[str, ...]


@dataclass(frozen=True)
class OpenQuestion:
    title: str
    choices: tuple[str, ...]


@dataclass(frozen=True)
class MatchingQuestion:
    title: str
    choices: tuple[tuple[str, str], ...]


REPLACEMENTS = {
    "\\": r"\textbackslash{}",
    "&": r"\&",
    "%": r"\%",
    "$": r"\$",
    "#": r"\#",
    "_": r"\_",
    "{": r"\{",
    "}": r"\}",
    "~": r"\textasciitilde{}",
    "^": r"\textasciicircum{}",
}


def escape_latex(string: str) -> str:
    return "".join(REPLACEMENTS.get(char, char) for char in string)


class WooclapConverter(GenericNodeVisitor):
    def __init__(self, document: document) -> None:
        super().__init__(document)
        self.output = ""
        self.latex = 0

    def visit_latex(self) -> None:
        if self.latex == 0:
            self.output += "$"

        self.latex += 1

    def depart_latex(self) -> None:
        self.latex -= 1

        if self.latex == 0:
            self.output += "$"

    def visit_document(self, node: document) -> None:
        self.output = ""

    def visit_literal_block(self, node: literal_block) -> Any:
        self.output += "```\n"

    def depart_literal_block(self, node: literal_block) -> Any:
        self.output += "\n```\n"

    def visit_strong(self, node: strong) -> Any:
        self.visit_latex()
        self.output += "\\textbf{"

    def depart_strong(self, node: strong) -> Any:
        self.output += "}"
        self.depart_latex()

    def visit_inline(self, node: inline) -> Any:
        pass

    def depart_inline(self, node: inline) -> Any:
        pass

    def visit_literal(self, node: literal) -> Any:
        self.visit_latex()
        self.output += "\\text{"

    def depart_literal(self, node: literal) -> Any:
        self.output += "}"
        self.depart_latex()

    def visit_system_message(self, node: system_message) -> Any:
        pass

    def visit_title_reference(self, node: title_reference) -> Any:
        pass

    def visit_paragraph(self, node: paragraph) -> Any:
        pass

    def depart_paragraph(self, node: paragraph) -> Any:
        self.output += "\n"

    def visit_Text(self, node: Text) -> Any:  # noqa: N802
        self.output += node if self.latex == 0 else escape_latex(node)

    def default_visit(self, node: Node) -> None:
        logger.warning(f"Unknown node type {type(node)}, skipping...")

    def default_departure(self, node: Node) -> None:
        pass


def convert_rst(source: str) -> str:
    doctree: document = publish_doctree(source)

    converter = WooclapConverter(doctree)

    doctree.walkabout(converter)

    return converter.output


def convert_title(problem: dict, always_title: bool) -> str:
    if "header" in problem:
        header = convert_rst(problem["header"])
        if always_title:
            return f"{problem['name']}\n\n{header}"
        else:
            return header
    else:
        return problem["name"]


def convert_multiple_choice_problem(problem: dict, always_title: bool) -> Qcm:
    title = convert_title(problem, always_title)

    correct: set[int] = set()
    choices: list[str] = []
    for i, choice in enumerate(problem["choices"]):
        choices.append(choice["text"])
        if choice["valid"]:
            correct.add(i + 1)

    return Qcm(title, frozenset(correct), tuple(choices))


def convert_regex_short_answer_problem(
    problem: dict, always_title: bool
) -> OpenQuestion:
    title = convert_title(problem, always_title)

    choices_ok = True
    choices: list[str] = []
    for match in problem["matches"]:
        regex: str = match["regex"]

        if match["valid"] and regex[0] == "^" and regex[-1] == "$":
            pattern = regex[1:-1]
            unescaped_pattern = pattern.replace("\\", "")
            if re.escape(unescaped_pattern) == pattern:
                choices.append(unescaped_pattern)
            else:
                choices_ok = False

    if not choices_ok:
        choices.clear()

    return OpenQuestion(title, tuple(choices))


def convert_matching_problem(problem: dict, always_title: bool) -> MatchingQuestion:
    title = convert_title(problem, always_title)

    choices: list[tuple[str, str]] = []
    for question in problem["questions"]:
        choices.append((convert_rst(question.get("question", "")), question["answer"]))

    return MatchingQuestion(title, tuple(choices))


def convert_match_problem(problem: dict, always_title: bool) -> OpenQuestion:
    title = convert_title(problem, always_title)

    return OpenQuestion(title, (problem["answer"],))


def convert_problem(
    problem: dict, always_title: bool
) -> Qcm | OpenQuestion | MatchingQuestion:
    match problem["type"]:
        case "multiple_choice":
            return convert_multiple_choice_problem(problem, always_title)
        case "regex_short_answer":
            return convert_regex_short_answer_problem(problem, always_title)
        case "matching":
            return convert_matching_problem(problem, always_title)
        case "match":
            return convert_match_problem(problem, always_title)
        case _:
            raise AssertionError("Unknown question type")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="The INGInious task to wooclap converter"
    )

    parser.add_argument("path", help="The path to the INGInious task file")
    parser.add_argument("output_path", help="The output path of the XLSX")
    parser.add_argument(
        "--title", action="store_true", help="Always show the title above the question"
    )

    args = parser.parse_args()

    path: str = args.path
    output_path: str = args.output_path
    always_title: bool = args.title

    with open(path, "r") as f:
        content: dict = yaml.safe_load(f)

    problems: dict[str, dict] = content["problems"]

    converted_problems = [
        convert_problem(problem, always_title) for problem in problems.values()
    ]

    max_choices = max(
        (len(problem.choices) for problem in converted_problems),
        default=0,
    )

    rows: list[tuple[str, ...]] = [
        ("Type", "Title", "Correct", *("Choice" for _ in range(max_choices)))
    ]

    for problem in converted_problems:
        match problem:
            case Qcm():
                rows.append(
                    (
                        "MCQ",
                        problem.title,
                        ",".join(str(id) for id in problem.correct),
                        *problem.choices,
                    )
                )
            case OpenQuestion():
                rows.append(
                    (
                        "OpenQuestion",
                        problem.title,
                        "",
                        *problem.choices,
                    )
                )
            case MatchingQuestion():
                rows.append(
                    (
                        "Matching",
                        problem.title,
                        "",
                        *(
                            f"{question} --- {answer}"
                            for (question, answer) in problem.choices
                        ),
                    )
                )

    workbook = xlsxwriter.Workbook(output_path)

    worksheet = workbook.add_worksheet()

    for i, row in enumerate(rows):
        for j, value in enumerate(row):
            worksheet.write(i, j, value)

    workbook.close()


if __name__ == "__main__":
    main()
