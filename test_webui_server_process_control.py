import asyncio
import unittest
from unittest.mock import AsyncMock, Mock, patch

import webui.server as mod


class _FakeAsyncStdout:
    def __init__(self, chunks):
        self._chunks = list(chunks)

    def __aiter__(self):
        self._iter = iter(self._chunks)
        return self

    async def __anext__(self):
        try:
            return next(self._iter)
        except StopIteration:
            raise StopAsyncIteration


class _FakeAsyncProcess:
    def __init__(self, chunks=None, pid=1234):
        self.stdout = _FakeAsyncStdout(chunks or [])
        self.returncode = None
        self.pid = pid
        self.send_signal = Mock()
        self.terminate = Mock()
        self.wait = AsyncMock(side_effect=self._finish)

    async def _finish(self):
        if self.returncode is None:
            self.returncode = 0
        return self.returncode


class WebuiServerProcessControlTests(unittest.IsolatedAsyncioTestCase):
    async def asyncSetUp(self):
        self.runs_backup = mod.RUNS
        mod.RUNS = {}

    async def asyncTearDown(self):
        mod.RUNS = self.runs_backup

    async def test_api_logs_keeps_streaming_after_ring_buffer_trim(self):
        rec = {"done": False, "lines": [], "line_offset": 0}
        mod.RUNS["r1"] = rec
        response = await mod.api_logs("r1")
        stream = response.body_iterator

        async def next_chunk():
            data = await asyncio.wait_for(anext(stream), timeout=1)
            return data.decode() if isinstance(data, bytes) else data

        rec["lines"] = [f"line{i}" for i in range(3)]
        self.assertIn("line0", await next_chunk())
        self.assertIn("line1", await next_chunk())
        self.assertIn("line2", await next_chunk())

        rec["line_offset"] = 2
        rec["lines"] = [f"line{i}" for i in range(2, 6)]
        self.assertIn("line3", await next_chunk())
        self.assertIn("line4", await next_chunk())
        self.assertIn("line5", await next_chunk())

        rec["done"] = True
        self.assertEqual("event: done\ndata: end\n\n", await next_chunk())


    async def test_api_stop_handles_asyncio_process(self):
        proc = _FakeAsyncProcess()
        mod.RUNS["r1"] = {"proc": proc, "done": False, "lines": []}

        with patch.object(mod, "stop_process_gracefully", AsyncMock()) as stopper:
            result = await mod.api_stop("r1")

        self.assertEqual({"ok": True}, result)
        stopper.assert_awaited_once_with(proc)


if __name__ == "__main__":
    unittest.main()
