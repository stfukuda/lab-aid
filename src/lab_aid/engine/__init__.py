"""Lab-Aid エンジンの公開インターフェースを提供するモジュール。"""

from __future__ import annotations

from .runtime import Engine, VarRef, evaluate

__all__ = ["Engine", "VarRef", "evaluate"]

if __name__ == "__main__":
    cases = [
        ("E", "this = #A * 2", "A=5"),
        ("R", "this = this + 3", "7"),
    ]
    for case in cases:
        result = evaluate(*case)
        print(case, "=>", result)
