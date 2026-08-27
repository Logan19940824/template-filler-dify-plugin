from __future__ import annotations

import argparse
import json
from pathlib import Path

from .services import fill_excel, fill_word, markdown_to_excel, markdown_to_word, parse_excel, parse_values, parse_word


def main() -> None:
    parser = argparse.ArgumentParser(prog="template-filler")
    sub = parser.add_subparsers(dest="command", required=True)
    for name in ("parse-word", "parse-excel"):
        command = sub.add_parser(name)
        command.add_argument("template", type=Path)
    for name in ("fill-word", "fill-excel"):
        command = sub.add_parser(name)
        command.add_argument("template", type=Path)
        command.add_argument("values", type=Path)
        command.add_argument("output", type=Path)
    for name in ("markdown-to-word", "markdown-to-excel"):
        command = sub.add_parser(name)
        command.add_argument("markdown", type=Path)
        command.add_argument("output", type=Path)
    args = parser.parse_args()
    if args.command == "parse-word":
        result = parse_word(args.template)
    elif args.command == "parse-excel":
        result = parse_excel(args.template)
    elif args.command in {"markdown-to-word", "markdown-to-excel"}:
        markdown_content = args.markdown.read_text(encoding="utf-8")
        converter = markdown_to_word if args.command == "markdown-to-word" else markdown_to_excel
        args.output.write_bytes(converter(markdown_content))
        return
    else:
        values = parse_values(json.loads(args.values.read_text(encoding="utf-8")))
        if args.command == "fill-word":
            fill_word(args.template, args.output, values)
        else:
            fill_excel(args.template, args.output, values)
        return
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
