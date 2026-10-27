"""Read a selected inline picture's camera from Word's in-memory range XML.

COM scalar getters remain primary. These finite Office presets and their default
X/Y angles were cross-checked with Word SetPresetCamera for all 62 enum values;
see the integration camera contract and native-camera-mapping-r1 evidence.
An omitted rot uses the preset, whereas an explicit rot overrides it.
"""
from hashlib import sha256
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
PIC = "{http://schemas.openxmlformats.org/drawingml/2006/picture}"
CAMERA_PROPERTIES = ("threed_preset_camera", "threed_rotation_x", "threed_rotation_y")

# XML preset: (MsoPresetCamera, default RotationX, default RotationY), degrees.
PRESET_CAMERAS = {
    "legacyObliqueTopLeft": (1, 0.0, 0.0),
    "legacyObliqueTop": (2, 0.0, 0.0),
    "legacyObliqueTopRight": (3, 0.0, 0.0),
    "legacyObliqueLeft": (4, 0.0, 0.0),
    "legacyObliqueFront": (5, 0.0, 0.0),
    "legacyObliqueRight": (6, 0.0, 0.0),
    "legacyObliqueBottomLeft": (7, 0.0, 0.0),
    "legacyObliqueBottom": (8, 0.0, 0.0),
    "legacyObliqueBottomRight": (9, 0.0, 0.0),
    "legacyPerspectiveTopLeft": (10, 0.0, 0.0),
    "legacyPerspectiveTop": (11, 0.0, 0.0),
    "legacyPerspectiveTopRight": (12, 0.0, 0.0),
    "legacyPerspectiveLeft": (13, 0.0, 0.0),
    "legacyPerspectiveFront": (14, 0.0, 0.0),
    "legacyPerspectiveRight": (15, 0.0, 0.0),
    "legacyPerspectiveBottomLeft": (16, 0.0, 0.0),
    "legacyPerspectiveBottom": (17, 0.0, 0.0),
    "legacyPerspectiveBottomRight": (18, 0.0, 0.0),
    "orthographicFront": (19, 0.0, 0.0),
    "isometricTopUp": (20, 45.2809, -35.3963),
    "isometricTopDown": (21, -45.2809, -35.3963),
    "isometricBottomUp": (22, -45.2809, 35.3963),
    "isometricBottomDown": (23, 45.2809, 35.3963),
    "isometricLeftUp": (24, -45.0, -35.0),
    "isometricLeftDown": (25, -45.0, 35.0),
    "isometricRightUp": (26, 45.0, 35.0),
    "isometricRightDown": (27, 45.0, -35.0),
    "isometricOffAxis1Left": (28, -64.0, 18.0),
    "isometricOffAxis1Right": (29, 26.0, 18.0),
    "isometricOffAxis1Top": (30, 53.4543, -58.7381),
    "isometricOffAxis2Left": (31, -26.0, 18.0),
    "isometricOffAxis2Right": (32, 64.0, 18.0),
    "isometricOffAxis2Top": (33, -53.4542, -58.7381),
    "isometricOffAxis3Left": (34, -64.0, -18.0),
    "isometricOffAxis3Right": (35, 26.0, -18.0),
    "isometricOffAxis3Bottom": (36, 53.4543, 58.7381),
    "isometricOffAxis4Left": (37, -26.0, -18.0),
    "isometricOffAxis4Right": (38, 64.0, -18.0),
    "isometricOffAxis4Bottom": (39, -53.4542, 58.7381),
    "obliqueTopLeft": (40, 0.0, 0.0),
    "obliqueTop": (41, 0.0, 0.0),
    "obliqueTopRight": (42, 0.0, 0.0),
    "obliqueLeft": (43, 0.0, 0.0),
    "obliqueRight": (44, 0.0, 0.0),
    "obliqueBottomLeft": (45, 0.0, 0.0),
    "obliqueBottom": (46, 0.0, 0.0),
    "obliqueBottomRight": (47, 0.0, 0.0),
    "perspectiveFront": (48, 0.0, 0.0),
    "perspectiveLeft": (49, -20.0, 0.0),
    "perspectiveRight": (50, 20.0, 0.0),
    "perspectiveAbove": (51, 0.0, -20.0),
    "perspectiveBelow": (52, 0.0, 20.0),
    "perspectiveAboveLeftFacing": (53, -14.3376, 39.3065),
    "perspectiveAboveRightFacing": (54, 14.3376, 39.3065),
    "perspectiveContrastingLeftFacing": (55, -43.9389, 10.3964),
    "perspectiveContrastingRightFacing": (56, 43.9389, 10.3964),
    "perspectiveHeroicLeftFacing": (57, -14.293, -11.0163),
    "perspectiveHeroicRightFacing": (58, 14.2931, -11.0163),
    "perspectiveHeroicExtremeLeftFacing": (59, -34.4607, 8.1224),
    "perspectiveHeroicExtremeRightFacing": (60, 34.4607, 8.1224),
    "perspectiveRelaxed": (61, 0.0, -50.4399),
    "perspectiveRelaxedModerately": (62, 0.0, -35.156),
}


def camera_from_range_xml(xml: str) -> dict:
    """Project only the single DrawingML picture owned by the selected range.

    Missing/ambiguous objects or unsupported camera states are unavailable, not
    evidence of an absent effect. Styles, themes and other package parts cannot
    provide the selected object's camera.
    """
    root = ET.fromstring(xml)
    bodies = list(root.iter(W + "body"))
    if len(bodies) != 1:
        raise ValueError("camera range must contain exactly one document body")
    body = bodies[0]
    pictures = list(body.iter(PIC + "pic"))
    drawings = list(body.iter(W + "drawing"))
    if len(pictures) != 1 or len(drawings) != 1:
        raise ValueError("camera range must own exactly one picture and drawing")
    owned = drawings[0].findall(
        f"{WP}inline/{A}graphic/{A}graphicData/{PIC}pic")
    if owned != pictures:
        raise ValueError("camera is not owned by the selected inline picture")
    scenes = pictures[0].findall(f"{PIC}spPr/{A}scene3d")
    if len(scenes) != 1:
        raise ValueError("selected picture must have exactly one explicit 3D scene")
    cameras = scenes[0].findall(A + "camera")
    if len(cameras) != 1:
        raise ValueError("selected picture must have exactly one explicit camera")
    camera = cameras[0]
    preset = camera.get("prst")
    if preset not in PRESET_CAMERAS:
        raise ValueError(f"unsupported camera preset: {preset}")
    enum, x, y = PRESET_CAMERAS[preset]
    children = list(camera)
    if children:
        if len(children) != 1 or children[0].tag != A + "rot":
            raise ValueError("unsupported camera rotation content")
        rot = children[0]
        if set(rot.attrib) != {"lat", "lon", "rev"} or list(rot):
            raise ValueError("camera rotation requires latitude, longitude and revolution")
        angles = {name: int(value) for name, value in rot.attrib.items()}
        if any(not 0 <= value < 21600000 for value in angles.values()):
            raise ValueError("camera rotation outside DrawingML angle range")
        # Word exposes signed angles: X is negative longitude, Y is latitude.
        x = ((-angles["lon"] / 60000 + 180) % 360) - 180
        y = ((angles["lat"] / 60000 + 180) % 360) - 180
    return {
        "threed_preset_camera": enum,
        "threed_rotation_x": round(x, 1),
        "threed_rotation_y": round(y, 1),
        "evidence": {
            "source": "word_com_range_wordopenxml",
            "range_xml_sha256": sha256(xml.encode("utf-8")).hexdigest(),
            "preset": preset,
            "rotation": dict(children[0].attrib) if children else None,
            "rotation_mode": "explicit" if children else "preset_default",
        },
    }


def read_inline_camera(inline) -> dict:
    """Capture the original one-character InlineShape range before conversion."""
    if inline.Type != 3:
        raise ValueError("camera fallback requires an inline picture")
    rng = inline.Range
    if rng.End - rng.Start != 1:
        raise ValueError("camera fallback requires the exact inline object range")
    return camera_from_range_xml(rng.WordOpenXML)


def recover_camera_properties(result, camera, error=None):
    """Fill failed getters only, keeping their original diagnostics auditable."""
    missing = {name: result.unavailable[name] for name in CAMERA_PROPERTIES
               if name in result.unavailable}
    if not missing:
        return
    if camera is None:
        result["_camera_xml_error"] = str(error)
        return
    # A readable COM value is authoritative. A disagreement means the two
    # projections cannot safely be combined on this Word build.
    for name in CAMERA_PROPERTIES:
        if name not in missing and result.get(name) != camera[name]:
            result["_camera_xml_error"] = f"native camera sources disagree on {name}"
            return
    result["_camera_xml_fallback"] = {
        **camera["evidence"], "com_read_errors": missing,
        "recovered_properties": sorted(missing),
    }
    for name in missing:
        result.read(name, lambda name=name: camera[name])
