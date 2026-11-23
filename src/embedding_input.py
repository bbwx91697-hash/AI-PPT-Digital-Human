"""Validate page selection before changing a PowerPoint copy."""


def validate_pages(pages, slide_count):
    seen = set()
    for page in pages:
        number = page["slide"]
        if not isinstance(number, int) or isinstance(number, bool) or not 1 <= number <= slide_count:
            raise ValueError(f"页码无效: {number}")
        if number in seen:
            raise ValueError(f"重复页码: {number}")
        seen.add(number)
