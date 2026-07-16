from markdown_it import MarkdownIt


class MarkdownParser:
    def __init__(self) -> None:
        self._parser = MarkdownIt("commonmark", {"html": False}).enable("table")

    def parse(self, source: str):
        return self._parser.parse(source)

    def parse_inline(self, source: str):
        return self._parser.parseInline(source)[0].children or []
