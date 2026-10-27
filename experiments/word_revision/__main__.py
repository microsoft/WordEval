"""Run the optional comparison through the shared evaluation runtime."""
import hashlib
from pathlib import Path


def main(argv=None):
    from scripts.evaluate_wordeval import main as evaluate
    import evaluation

    root = Path(evaluation.__file__).resolve().parent
    sources = [root / '__init__.py', root / 'evaluator.py', root / 'isolated.py', root / 'isolated_worker.py',
               root / 'methods/__init__.py',
               root / 'methods/word_revision.py', root / 'methods/base.py',
               *sorted((root / 'document_io').rglob('*.py'))]
    metadata = {
        'name': 'word_revision',
        'protocol': 'normal-open-shared-runtime-v1',
        'role': 'optional_comparison',
        'source_sha256': {p.relative_to(root).as_posix(): hashlib.sha256(p.read_bytes()).hexdigest()
                          for p in sources},
        'entrypoint_sha256': {
            p.relative_to(root.parent).as_posix(): hashlib.sha256(p.read_bytes()).hexdigest()
            for p in (root.parent / 'scripts/evaluate_wordeval.py', Path(__file__).resolve())
        },
    }
    return evaluate(argv, default_methods=['word_revision'], allowed_methods=['word_revision'],
                    experiment_metadata=metadata)


if __name__ == '__main__':
    raise SystemExit(main())
