"""Canonical effective character intervals and preservation evidence checks."""
from ...observations import ObservedProperties, require_observations
from ....base import VerificationResult
from ..references.fields import INSTRUCTION_ONLY_FIELD_TYPES


INACTIVE_CHARACTER_FIELDS = {
    "text_outline_color", "text_outline_weight", "text_outline_dash_style", "text_fill_color",
}


def text_width(text):
    return len(text.encode("utf-16-le", errors="surrogatepass")) // 2


def character_projection(rng):
    """Map native Word characters onto the text returned by Range.Text.

    Content-control boundaries occupy native positions without producing text.
    A field occupies many native positions but one Characters collection item;
    only its Result contributes visible formatting, not its instruction code.
    Plain ranges need no projection or per-character COM calls.
    """
    text = rng.Text
    if rng.End - rng.Start == text_width(text):
        return None
    fields = {}
    for index in range(1, rng.Fields.Count + 1):
        field = rng.Fields(index)
        if field.Type in INSTRUCTION_ONLY_FIELD_TYPES:
            continue
        result = field.Result.Duplicate
        span = (int(field.Code.Start) - 1, int(result.End) + 1)
        if span in fields:
            raise ValueError("Duplicate native field span")
        fields[span] = result
    projected, observed = [], []
    characters = rng.Characters
    for index in range(1, characters.Count + 1):
        character = characters(index)
        # Word can return a character spanning a field delimiter or an anchor
        # outside this paragraph. Only its intersection with the owned range
        # contributes text or formatting.
        start, end = max(character.Start, rng.Start), min(character.End, rng.End)
        if start >= end:
            continue
        owned = character.Duplicate
        owned.Start, owned.End = start, end
        content = owned.Text
        observed.append(content)
        if not content:
            continue
        # A multi-paragraph field can begin in this paragraph and end in a
        # later one. Characters exposes only the local part of that field.
        containers = [(right - left, result) for (left, right), result in fields.items()
                      if left <= start and end <= right]
        if containers:
            target = min(containers, key=lambda item: item[0])[1].Duplicate
            start, end = max(start, target.Start), min(end, target.End)
            if start >= end:
                raise ValueError("Field result does not cover its native character text")
            target.Start, target.End = start, end
        else:
            target = owned
        segments = [target]
        if target.Text != content:
            # Some native fields expose a delimiter in Range.Text outside
            # Result (for example a legacy SHAPE field's trailing U+0015).
            # Read that delimiter at its own position; never import field
            # instruction formatting or silently discard observed text.
            prefix, suffix = owned.Duplicate, owned.Duplicate
            prefix.End, suffix.Start = target.Start, target.End
            for boundary in (prefix, suffix):
                value = boundary.Text
                if value and (set(value) - {"\x13", "\x14", "\x15"}
                              or boundary.End - boundary.Start != text_width(value)):
                    raise ValueError("Field result differs from its native character text")
            segments = [part for part in (prefix, target, suffix) if part.Text]
            if "".join(part.Text for part in segments) != content:
                raise ValueError("Field result differs from its native character text")
        projected.extend((part, text_width(part.Text)) for part in segments)
    if "".join(observed) != text:
        raise ValueError("Native character collection does not cover the observed text")
    return projected


def _merge_intervals(values):
    intervals = []
    for start, end, value in values:
        if intervals and intervals[-1][1] == start and intervals[-1][2] == value:
            intervals[-1] = (intervals[-1][0], end, value)
        else:
            intervals.append((start, end, value))
    return intervals[0][2] if len(intervals) == 1 else intervals


_UNSET = object()


def character_property(doc, rng, getter, *, normalize=lambda value: value,
                       empty_is_mixed=False, nullable=False, office_mixed=False,
                       projection=_UNSET):
    """Split mixed ranges, then merge adjacent equal intervals.

    Uniform paragraphs need one getter call. Mixed paragraphs are compared by
    character position, independent of the run boundaries chosen when saved.
    """
    if projection is _UNSET:
        projection = character_projection(rng)
    if projection is not None:
        intervals, cursor = [], 0
        for segment, width in projection:
            # A native end-of-cell character exposes two text characters
            # (CR + BEL), but occupies one position and cannot be subdivided.
            same_span = (segment.Start, segment.End) == (rng.Start, rng.End)
            value = character_property(doc, segment, getter, normalize=normalize,
                empty_is_mixed=empty_is_mixed, nullable=nullable, office_mixed=office_mixed,
                projection=None if same_span else _UNSET)
            if same_span and segment.End - segment.Start != width and isinstance(value, list):
                raise ValueError("Native character has ambiguous visible formatting")
            if isinstance(value, list):
                intervals.extend((cursor + start, cursor + end, item) for start, end, item in value)
            else:
                intervals.append((cursor, cursor + width, value))
            cursor += width
        return _merge_intervals(intervals)
    origin = rng.Start

    def collect(segment):
        value = getter(segment)
        mixed = value in (9999999, -9999999) or (office_mixed and value == -2)
        mixed = mixed or (empty_is_mixed and value == "")
        if mixed:
            if segment.End - segment.Start <= 1:
                raise ValueError("Character property remains mixed on a single character")
            middle = (segment.Start + segment.End) // 2
            left, right = segment.Duplicate, segment.Duplicate
            left.End, right.Start = middle, middle
            return collect(left) + collect(right)
        state = ObservedProperties()
        state.read("value", lambda: value, normalize)
        error = require_observations({"range": state}, ["value"], context="character property",
                                     nullable=["value"] if nullable else [])
        if error is not None:
            raise ValueError(error.details)
        return [(segment.Start - origin, segment.End - origin, state["value"])]

    return _merge_intervals(collect(rng))


def validate_character_observations(states, *, ignored=()):
    """Validate full reader projections before dict equality can lose metadata."""
    for owner, paragraphs in states.items():
        if paragraphs is None or (owner == "gold" and not paragraphs):
            return VerificationResult.error_result("Character projection is unavailable", observation_owner=owner)
        for index, state in enumerate(paragraphs, 1):
            fields = set(state) - set(ignored)
            error = require_observations({owner: state}, fields,
                context=f"paragraph {index} character formatting preservation", nullable=INACTIVE_CHARACTER_FIELDS)
            if error is not None:
                return error
    return None


def visible_character_stream(paragraphs):
    """Compare effective formatting across permitted paragraph/page splits."""
    stream = []
    for paragraph in paragraphs:
        offset = 0
        for character in paragraph.get("text", ""):
            properties = {}
            for name, value in paragraph.items():
                if name in {"index", "text", "hidden_ranges"}:
                    continue
                if isinstance(value, list):
                    matches = [item[2] for item in value if item[0] <= offset < item[1]]
                    if len(matches) != 1:
                        raise ValueError(f"Incomplete character interval for {name} at {offset}")
                    value = matches[0]
                properties[name] = value
            # Word range offsets count UTF-16 code units.
            offset += len(character.encode("utf-16-le")) // 2
            if character in {"\r", "\x0c", "\x07"}:
                continue
            if stream and stream[-1]["format"] == properties:
                stream[-1]["text"] += character
            else:
                stream.append({"text": character, "format": properties})
    return stream
