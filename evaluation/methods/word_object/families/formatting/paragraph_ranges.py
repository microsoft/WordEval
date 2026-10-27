"""Native paragraph ranges, excluding Word's synthetic table-row marker."""


def is_row_end(para):
    rng = para.Range
    if rng.Text != "\r\x07":
        return False
    start = rng.Duplicate
    start.Collapse(1)  # wdCollapseStart
    return bool(start.IsEndOfRowMark)


def paragraph_text_range(para):
    rng = para.Range.Duplicate
    while rng.End > rng.Start:
        tail = rng.Duplicate
        tail.Start = tail.End - 1
        if tail.Text not in {"\r", "\x07", "\r\x07"}:
            break
        rng.End -= 1
    return rng
