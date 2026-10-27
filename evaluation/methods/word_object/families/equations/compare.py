"""families / equations / compare with explicit evidence dependencies."""

from typing import Any, Dict
from copy import deepcopy
import re
import unicodedata


def normalized_equation_text(equation: Dict[str, Any]) -> str:
    superscripts = str.maketrans(
        {
            "⁰": "^(0)",
            "¹": "^(1)",
            "²": "^(2)",
            "³": "^(3)",
            "⁴": "^(4)",
            "⁵": "^(5)",
            "⁶": "^(6)",
            "⁷": "^(7)",
            "⁸": "^(8)",
            "⁹": "^(9)",
            "⁺": "^(+)",
            "⁻": "^(-)",
        }
    )
    translated = (equation.get("equation_text") or "").strip().translate(superscripts)
    normalized = unicodedata.normalize("NFKC", translated)
    return "".join(normalized.split())


def folded_equation_text(equation: Dict[str, Any]) -> str:
    normalized = unicodedata.normalize("NFKC", (equation.get("equation_text") or "").strip())
    return "".join(normalized.split())


def normalized_basic_operator_text(equation: Dict[str, Any]) -> str:
    """Normalize Word-native spellings of text-only basic operators."""
    normalized = normalized_equation_text(equation)
    return normalized.replace("\u2236=", ":=").replace("\u2254", ":=")


def canonical_angle_bracket(value: Any) -> Any:
    return {
        0x2329: 0x27E8,
        0x232A: 0x27E9,
    }.get(value, value)


def normalized_angle_delimiter_text(equation: Dict[str, Any]) -> str:
    """Normalize only Word's legacy code points for the angle gallery.

    The replacement happens before NFKC because NFKC itself maps U+2329
    and U+232A to the CJK pair U+3008/U+3009.  Replacing first preserves
    the approved U+2329/U+232A <-> U+27E8/U+27E9 boundary without also
    accepting literal CJK delimiters.
    """
    text = (equation.get("equation_text") or "").strip()
    text = text.translate({0x2329: 0x27E8, 0x232A: 0x27E9})
    normalized = unicodedata.normalize("NFKC", text)
    return "".join(normalized.split())


def has_native_angle_delimiter(equation: Dict[str, Any]) -> bool:
    """Return whether WOM reports a native angle-delimiter function."""
    for function in equation.get("functions", []):
        if function.get("type") != 5:  # wdOMathFunctionDelim
            continue
        begin = canonical_angle_bracket(function.get("delim_begin_char"))
        end = canonical_angle_bracket(function.get("delim_end_char"))
        if (begin, end) == (0x27E8, 0x27E9):
            return True
    return False


def equation_com_semantic_signature(equation: Dict[str, Any]) -> Dict[str, Any]:
    semantic_keys = {
        "type",
        "accent_char",
        "bar_top",
        "delim_begin_char",
        "delim_end_char",
        "delim_sep_char",
        "frac_type",
        "mat_rows",
        "mat_cols",
        "nary_char",
        "nary_sub_sup_lim",
        "nary_hide_sub",
        "nary_hide_sup",
        "rad_hide_degree",
        "group_char",
        "group_char_top",
        "semantic_expression",
        "argument_expressions",
    }

    def semantic_function(function: Dict[str, Any]) -> Dict[str, Any]:
        result = {key: function.get(key) for key in semantic_keys if key in function}
        if function.get("nary_hide_sub") and function.get("nary_hide_sup"):
            result.pop("nary_sub_sup_lim", None)
        for key, value in function.items():
            if key.endswith("_nested") and isinstance(value, list):
                result[key] = [semantic_function(item) for item in value]
        return result

    return {
        "equation_count": equation.get("equation_count"),
        "equation_text": normalized_equation_text(equation),
        "functions": [semantic_function(function) for function in equation.get("functions", [])],
        **({"equations": [equation_com_semantic_signature(item)
                          for item in equation['equations']]}
           if 'equations' in equation else {}),
    }


def compact_cases_com_signature(equation: Dict[str, Any]) -> Dict[str, Any]:
    """Apply the named three-condition contract to equation 1 only.

    A compact 3x1 matrix and EqArray can encode the same cases. Normalize
    their container/row-role spelling, retaining every typed row child.
    Callers must separately require the saved alignment/spacing guardrail.
    Other equations and nested matrices are preservation evidence.
    """
    signature = equation_com_semantic_signature(equation)

    def normalize_target(state):
        for delimiter in state.get("functions", []):
            children = delimiter.get("delim_1_nested", [])
            if (delimiter.get("type") != 5 or len(children) != 1
                    or delimiter.get("delim_begin_char") != ord("{")
                    or delimiter.get("delim_end_char") != 0):
                continue
            matrix = children[0]
            roles = [f"{row},1" for row in range(1, 4)]
            nested = [f"mat_{row}_1_nested" for row in range(1, 4)]
            arguments = matrix.get("argument_expressions", {})
            if (matrix.get("type") != 12 or matrix.get("mat_rows") != 3
                    or matrix.get("mat_cols") != 1 or set(arguments) != set(roles)
                    or {key for key in matrix if key.endswith("_nested")} != set(nested)
                    or not all(isinstance(matrix[key], list) for key in nested)):
                continue
            old_expression = "Mat(" + ";".join(
                f"{role}={arguments[role]}" for role in roles
            ) + ")"
            if (matrix.get("semantic_expression") != old_expression
                    or delimiter.get("argument_expressions") != {"1": old_expression}
                    or delimiter.get("semantic_expression") != "{" + old_expression + "\0"):
                continue
            new_arguments = {str(row): arguments[f"{row},1"] for row in range(1, 4)}
            new_expression = "EqArray(" + ";".join(
                f"{role}={value}" for role, value in new_arguments.items()
            ) + ")"
            matrix.update(type=6, argument_expressions=new_arguments,
                          semantic_expression=new_expression)
            matrix.pop("mat_rows")
            matrix.pop("mat_cols")
            for row, key in enumerate(nested, start=1):
                matrix[f"eqarray_{row}_nested"] = matrix.pop(key)
            delimiter["argument_expressions"] = {"1": new_expression}
            delimiter["semantic_expression"] = "{" + new_expression + "\0"

    normalize_target(signature)
    if signature.get("equations"):
        normalize_target(signature["equations"][0])
    return signature


def linear_conversion_com_signature(equation: Dict[str, Any]) -> Dict[str, Any]:
    """Normalize isolated digit exponents only in a single MathText linear run.

    Word BuildUp/Linearize maps both x^2 and x² to the same native superscript.
    Keep multi-digit/compound exponents and typed structures distinct. The
    caller restricts this equivalence to equations targeted by conversion.
    """
    signature = equation_com_semantic_signature(equation)
    functions = signature["functions"]
    if (equation.get("function_count") == 1 and len(functions) == 1
            and functions[0].get("type") == 20):
        def exponent_text(text):
            return re.sub(r"\^([0-9])(?=$|[+\-=<>*/)])", r"^(\1)", text)

        signature["equation_text"] = exponent_text(signature["equation_text"])
        if "semantic_expression" in functions[0]:
            functions[0]["semantic_expression"] = exponent_text(functions[0]["semantic_expression"])
    return signature


def logarithm_name_com_signature(equation: Dict[str, Any]) -> Dict[str, Any]:
    """Permit upright log operator text in the target native FName role.

    The named logarithm contract requires log, its base, and its argument;
    it does not request italic operator lettering. This is not NormalText
    equivalence for variables or for any subsequent preservation equation.
    """
    signature = equation_com_semantic_signature(equation)

    def normalize_target(state):
        for function in state.get("functions", []):
            names = function.get("func_fname_nested", [])
            if function.get("type") != 8 or len(names) != 1:
                continue
            name = names[0]
            text = name.get("scrsub_nested", [])
            if (name.get("type") == 17 and len(text) == 1
                    and name.get("argument_expressions", {}).get("E") == "log"
                    and text[0].get("type") == 21
                    and text[0].get("semantic_expression") == "log"):
                text[0]["type"] = 20

    normalize_target(signature)
    if signature.get("equations"):
        normalize_target(signature["equations"][0])
    return signature


def equation_contract_equivalence(spec: Dict[str, Any]) -> str:
    """Return only the explicitly approved, named contract equivalences."""
    if (spec.get("idMso") == "ScopeChangingEquation"
            and spec.get("variation") == "variant_fraction_integral"
            and spec.get("composition", {}).get("interaction_class") == "insert_fraction_then_integral"):
        return "fraction_integral_limit_placement"
    if (spec.get("idMso") == "EquationMatrixGallery"
            and spec.get("variation") == "sparse_3x3_brackets"):
        return "sparse_matrix_index_notation"
    return ""


def normalize_equation_contract(equation: Dict[str, Any], equivalence: str) -> Dict[str, Any]:
    """Normalize a comparison copy, preserving the complete native observation.

    Only equation 1 is targeted. Other equations and every unapproved field
    remain literal preservation evidence. The caller retains the raw evidence.
    """
    result = deepcopy(equation)
    for state in [result] + result.get("equations", [])[:1]:
        functions = state.get("functions", [])
        if len(functions) != 2:
            continue
        target = functions[1]
        if equivalence == "fraction_integral_limit_placement":
            children = target.get("frac_num_nested", [])
            if target.get("type") != 7 or len(children) != 1:
                continue
            integral = children[0]
            if (integral.get("type") == 13 and integral.get("nary_char") == 8747
                    and integral.get("nary_sub_sup_lim") in (0, -1, True)):
                integral["nary_sub_sup_lim"] = 0
        elif equivalence == "sparse_matrix_index_notation":
            _normalize_sparse_matrix_indices(target)
    return result


def _normalize_sparse_matrix_indices(delimiter: Dict[str, Any]) -> None:
    """Collapse approved native a-subscript cells to their plain label form.

    Require complete typed leaves and consistent cached expressions before
    updating the matrix and delimiter expressions. Do not flatten arbitrary
    subscripts, extra child roles, or subscripts in other matrix positions.
    """
    matrices = delimiter.get("delim_1_nested", [])
    if (delimiter.get("type") != 5 or delimiter.get("delim_begin_char") != 91
            or delimiter.get("delim_end_char") != 93 or len(matrices) != 1):
        return
    matrix = matrices[0]
    roles = [f"{row},{column}" for row in range(1, 4) for column in range(1, 4)]
    arguments = matrix.get("argument_expressions", {})
    if (matrix.get("type") != 12 or matrix.get("mat_rows") != 3
            or matrix.get("mat_cols") != 3 or set(arguments) != set(roles)):
        return
    old_expression = "Mat(" + ";".join(f"{role}={arguments[role]}" for role in roles) + ")"
    if (matrix.get("semantic_expression") != old_expression
            or delimiter.get("argument_expressions") != {"1": old_expression}
            or delimiter.get("semantic_expression") != "[" + old_expression + "]"):
        return

    def leaf(value):
        return {"index": 1, "type": 20, "semantic_expression": value}

    for role, index in (("1,1", "11"), ("1,3", "1n"), ("3,1", "n1"), ("3,3", "nn")):
        key = "mat_" + role.replace(",", "_") + "_nested"
        expression = f"a_({index})"
        expected = {"index": 1, "type": 17, "scrsub_nested": [leaf("a")],
                    "scrsub_sub_nested": [leaf(index)],
                    "argument_expressions": {"E": "a", "Sub": index},
                    "semantic_expression": expression}
        if matrix.get(key) == [expected] and arguments[role] == expression:
            matrix[key] = [leaf("a" + index)]
            arguments[role] = "a" + index
    expression = "Mat(" + ";".join(f"{role}={arguments[role]}" for role in roles) + ")"
    matrix["semantic_expression"] = expression
    delimiter["argument_expressions"] = {"1": expression}
    delimiter["semantic_expression"] = "[" + expression + "]"
