from lab_aid import hello


def test_hello_returns_greeting() -> None:
    assert hello("Codex") == "Hello, Codex!"
