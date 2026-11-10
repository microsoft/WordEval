/* Copyright (c) Microsoft Corporation. Licensed under the MIT License. */
/* Failure analysis from the paper build (WordEval-r3): failure_summary.csv and failure_forms.csv.
   Regenerate rather than edit by hand. */
window.WORDEVAL_FAILURES = {
  "package": "WordEval-r3",
  "n": 642,
  "systems": [
    {
      "id": "claude-code-opus-5",
      "passed": 559,
      "incorrect": 59,
      "invalid": 0,
      "missing": 24
    },
    {
      "id": "codex-gpt-56-sol-high",
      "passed": 428,
      "incorrect": 96,
      "invalid": 94,
      "missing": 24
    },
    {
      "id": "kimi-code-k3-high",
      "passed": 397,
      "incorrect": 143,
      "invalid": 13,
      "missing": 89
    },
    {
      "id": "mnm_minimax_m3",
      "passed": 286,
      "incorrect": 245,
      "invalid": 44,
      "missing": 67
    },
    {
      "id": "qwen-code-qwen38-27b",
      "passed": 305,
      "incorrect": 214,
      "invalid": 49,
      "missing": 74
    },
    {
      "id": "opencode-gemini-38-flash-high",
      "passed": 363,
      "incorrect": 11,
      "invalid": 0,
      "missing": 268
    },
    {
      "id": "zcode-glm-52-high",
      "passed": 356,
      "incorrect": 179,
      "invalid": 49,
      "missing": 58
    },
    {
      "id": "opencode-deepseek-v4-pro",
      "passed": 339,
      "incorrect": 250,
      "invalid": 44,
      "missing": 9
    }
  ],
  "mismatch": {
    "id": "claude-code-opus-5",
    "n": 59,
    "types": [
      {
        "key": "property_value_mismatch",
        "n": 23
      },
      {
        "key": "structure_mismatch",
        "n": 20
      },
      {
        "key": "missing_element",
        "n": 7
      },
      {
        "key": "off_target_change",
        "n": 4
      },
      {
        "key": "content_mismatch",
        "n": 4
      },
      {
        "key": "unrequested_change",
        "n": 1
      }
    ]
  }
};
