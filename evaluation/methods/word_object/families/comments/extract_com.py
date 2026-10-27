"""families / comments / extract_com with explicit evidence dependencies."""

from typing import Dict


def extract_comment_info(docx_path: str, *, close_document, open_document) -> Dict:
    """
    Extract detailed comment information from document.

    Returns:
        Dict with:
        - comment_count: Total number of comments
        - comments: List of comment details including:
            - index, text, scope_text, author, date
            - is_reply: Whether it's a reply (has Ancestor)
            - is_resolved: Whether it's marked as Done
            - ancestor_index: Parent comment index (for replies)
    """
    doc = None
    try:
        doc = open_document(docx_path)

        result = {
            "comment_count": 0,
            "top_level_count": 0,
            "comments": [],
            "resolved_count": 0,
            "reply_count": 0,
        }

        comments = doc.Comments
        seen_dispatches = set()
        seen_comment_locations = set()

        def dispatch_identity(comment):
            """Return a stable in-process COM identity when pywin32 exposes it."""
            try:
                return int(comment._oleobj_.GetIUnknown())
            except Exception:
                return None

        def range_details(range_object):
            if range_object is None:
                return "", None, None
            return (
                (range_object.Text or "").strip(),
                int(range_object.Start),
                int(range_object.End),
            )

        def append_comment(
            comment,
            *,
            parent_info=None,
            merge_existing_location=False,
        ):
            identity = dispatch_identity(comment)
            if identity is not None and identity in seen_dispatches:
                return None
            if identity is not None:
                seen_dispatches.add(identity)

            ancestor = None
            if parent_info is None:
                try:
                    ancestor = comment.Ancestor
                except Exception:
                    # Word raises for a top-level comment instead of returning
                    # None on some builds.
                    ancestor = None

            is_reply = parent_info is not None or ancestor is not None
            text, range_start, range_end = range_details(comment.Range)
            try:
                scope_text, scope_start, scope_end = range_details(comment.Scope)
            except Exception:
                if parent_info is not None:
                    scope_text = parent_info.get("scope_text", "")
                    scope_start = parent_info.get("scope_start")
                    scope_end = parent_info.get("scope_end")
                elif ancestor is not None:
                    scope_text, scope_start, scope_end = range_details(ancestor.Scope)
                else:
                    raise
            ancestor_index = None
            ancestor_range_start = None
            if parent_info is not None:
                ancestor_index = parent_info["index"]
                ancestor_range_start = parent_info.get("range_start")
                # Replies frequently expose an empty Scope through COM; the
                # semantic scope is inherited from the parent thread.
                if not scope_text:
                    scope_text = parent_info.get("scope_text", "")
                    scope_start = parent_info.get("scope_start")
                    scope_end = parent_info.get("scope_end")
            elif ancestor is not None:
                try:
                    ancestor_index = int(ancestor.Index)
                except Exception:
                    ancestor_index = None
                try:
                    ancestor_range_start = int(ancestor.Range.Start)
                except Exception:
                    ancestor_range_start = None

            try:
                is_resolved = bool(comment.Done)
            except Exception:
                # Done is unavailable for replies and on older Word object
                # models. Resolution is a top-level-thread property.
                is_resolved = False

            author = comment.Author or ""
            location_signature = (
                range_start,
                range_end,
                text,
                author,
                ancestor_range_start if is_reply else None,
            )
            if merge_existing_location and location_signature in seen_comment_locations:
                return None
            seen_comment_locations.add(location_signature)

            info = {
                "index": len(result["comments"]) + 1,
                "collection_index": int(comment.Index),
                "text": text,
                "scope_text": scope_text,
                "author": author,
                "is_reply": is_reply,
                "is_resolved": is_resolved,
                "ancestor_index": ancestor_index,
                "ancestor_range_start": ancestor_range_start,
                "range_start": range_start,
                "range_end": range_end,
                "scope_start": scope_start,
                "scope_end": scope_end,
            }
            result["comments"].append(info)
            result["reply_count"] += int(is_reply)
            result["resolved_count"] += int(is_resolved)
            return info

        top_level_infos = []
        for i in range(1, int(comments.Count) + 1):
            comment = comments(i)
            info = append_comment(comment)
            if info is not None and not info["is_reply"]:
                top_level_infos.append((comment, info))

        # Modern threaded replies are not consistently included in
        # Document.Comments. Traverse each top-level Replies collection and
        # merge by COM identity when Word exposed the reply twice.
        for parent, parent_info in top_level_infos:
            try:
                replies = parent.Replies
                reply_count = int(replies.Count)
            except Exception:
                continue
            for reply_index in range(1, reply_count + 1):
                append_comment(
                    replies(reply_index),
                    parent_info=parent_info,
                    merge_existing_location=True,
                )

        result["comment_count"] = len(result["comments"])
        result["top_level_count"] = sum(
            1 for comment in result["comments"] if not comment["is_reply"]
        )
        return result

    finally:
        close_document(doc)
