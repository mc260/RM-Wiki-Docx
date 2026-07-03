"""CLI entry point for wiki2docx."""

import argparse
import json
import os
import sys

from .scraper import WikiScraper
from .converter import build_docx


def main():
    parser = argparse.ArgumentParser(
        prog="wiki2docx",
        description="将 RoboMaster 论坛 Wiki 文档导出为 Word (.docx) 文件",
    )
    parser.add_argument(
        "--wiki-id",
        type=str,
        required=False,
        default=None,
        help="Wiki 页面 ID（例如 46368458）。使用 --json-file 时可不提供。",
    )
    parser.add_argument(
        "--output",
        type=str,
        default="output.docx",
        help="输出 Word 文件路径（默认: output.docx）",
    )
    parser.add_argument(
        "--config",
        type=str,
        default=None,
        help="配置文件路径（JSON 格式）",
    )
    parser.add_argument(
        "--profile-dir",
        type=str,
        default=None,
        help="工作目录（默认: ./_browser_profile），用于存放截图和日志",
    )
    parser.add_argument(
        "--json-only",
        action="store_true",
        help="仅抓取并保存 JSON，不生成 Word 文档",
    )
    parser.add_argument(
        "--json-file",
        type=str,
        default=None,
        help="跳过抓取，直接从已有的 JSON 文件生成 Word",
    )

    args = parser.parse_args()

    # Load config
    config = {}
    if args.config:
        if not os.path.exists(args.config):
            print(f"错误: 配置文件不存在: {args.config}")
            sys.exit(1)
        with open(args.config, "r", encoding="utf-8") as f:
            config = json.load(f)
    else:
        # Try default config location
        default_config = os.path.join(os.path.dirname(__file__), "..", "config.json")
        if os.path.exists(default_config):
            with open(default_config, "r", encoding="utf-8") as f:
                config = json.load(f)

    wiki_url = config.get("wiki_url", f"https://bbs.robomaster.com/wiki/{args.wiki_id}")
    title_override = config.get("title")
    subtitle_override = config.get("subtitle")
    font_name = config.get("font", "微软雅黑")

    data = None

    # --- Validate ---
    if not args.json_file and not args.wiki_id:
        print("错误: 必须提供 --wiki-id 或 --json-file 参数。")
        print("示例: python -m wiki2docx --wiki-id 46368458")
        print("示例: python -m wiki2docx --json-file data.json --output output.docx")
        sys.exit(1)

    # --- JSON file mode: skip scraping ---
    if args.json_file:
        if not os.path.exists(args.json_file):
            print(f"错误: JSON 文件不存在: {args.json_file}")
            sys.exit(1)
        print(f"从已有 JSON 文件生成: {args.json_file}")
        with open(args.json_file, "r", encoding="utf-8") as f:
            data = json.load(f)

    # --- Scrape mode ---
    else:
        profile_dir = args.profile_dir or os.path.join(
            os.path.dirname(__file__), "..", "_browser_profile"
        )
        profile_dir = os.path.abspath(profile_dir)

        scraper = WikiScraper(
            wiki_id=args.wiki_id,
            profile_dir=profile_dir,
        )
        data = scraper.scrape()

        # Save intermediate JSON
        json_path = args.output.replace(".docx", ".json")
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"JSON 已保存: {json_path}")

    if args.json_only:
        print("--json-only 模式，跳过 Word 生成")
        return

    # --- Build DOCX ---
    output_path = os.path.abspath(args.output)
    build_docx(
        data=data,
        output_path=output_path,
        title_override=title_override,
        subtitle_override=subtitle_override,
        font_name=font_name,
    )
    print(f"Word 文档已生成: {output_path}")


if __name__ == "__main__":
    main()
