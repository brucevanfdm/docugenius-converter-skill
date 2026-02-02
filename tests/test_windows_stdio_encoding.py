import io
import os
import sys
import unittest


class _TtyWrapper:
    def __init__(self, inner):
        self._inner = inner

    def isatty(self):
        return True

    def write(self, s):
        return self._inner.write(s)

    def flush(self):
        return self._inner.flush()

    @property
    def encoding(self):
        return getattr(self._inner, "encoding", None)

    @property
    def errors(self):
        return getattr(self._inner, "errors", None)

    @property
    def buffer(self):
        return getattr(self._inner, "buffer", None)

    def detach(self):
        return self._inner.detach()

    def reconfigure(self, **kwargs):
        return self._inner.reconfigure(**kwargs)


class TestWindowsStdIOEncoding(unittest.TestCase):
    @unittest.skipUnless(sys.platform == "win32", "仅在 Windows 上运行")
    def test_non_tty_stdout_gbk_strict_does_not_crash(self):
        sys.path.insert(0, os.path.abspath("scripts"))
        import convert_document  # noqa: E402

        old_stdout, old_stderr = sys.stdout, sys.stderr
        try:
            sys.stdout = io.TextIOWrapper(io.BytesIO(), encoding="gbk", errors="strict")
            sys.stderr = io.TextIOWrapper(io.BytesIO(), encoding="gbk", errors="strict")
            convert_document._configure_windows_stdio()
            print("✓")  # 不应抛 UnicodeEncodeError
        finally:
            sys.stdout, sys.stderr = old_stdout, old_stderr

    @unittest.skipUnless(sys.platform == "win32", "仅在 Windows 上运行")
    def test_tty_stdout_gbk_strict_does_not_crash(self):
        sys.path.insert(0, os.path.abspath("scripts"))
        import convert_document  # noqa: E402

        old_stdout, old_stderr = sys.stdout, sys.stderr
        try:
            inner_out = io.TextIOWrapper(io.BytesIO(), encoding="gbk", errors="strict")
            inner_err = io.TextIOWrapper(io.BytesIO(), encoding="gbk", errors="strict")
            sys.stdout = _TtyWrapper(inner_out)
            sys.stderr = _TtyWrapper(inner_err)
            convert_document._configure_windows_stdio()
            print("✗")  # 不应抛 UnicodeEncodeError
        finally:
            sys.stdout, sys.stderr = old_stdout, old_stderr


if __name__ == "__main__":
    unittest.main()

