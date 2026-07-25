import unittest
from unittest.mock import patch

import register_outlook_ruoyi as mod


class _FakeClock:
    def __init__(self, start=100.0):
        self.now = float(start)

    def time(self):
        return self.now

    def sleep(self, seconds):
        self.now += max(0.0, float(seconds))


class _FakeActions:
    def __init__(self, clock):
        self.clock = clock
        self.calls = []

    def move_to(self, loc, duration=None):
        self.calls.append(("move_to", loc, duration))
        return self

    def hold(self):
        self.calls.append(("hold",))
        return self

    def wait(self, seconds):
        self.calls.append(("wait", seconds))
        self.clock.sleep(seconds)
        return self

    def release(self):
        self.calls.append(("release",))
        return self

    def perform(self):
        self.calls.append(("perform",))
        return self

    def release_all(self):
        self.calls.append(("release_all",))
        return self


class _FakeContext:
    def __init__(self, clock):
        self.actions = _FakeActions(clock)


class PxHoldEarlyReleaseTests(unittest.TestCase):
    def test_hold_releases_early_when_hold_label_turns_display_none(self):
        clock = _FakeClock()
        ctx = _FakeContext(clock)
        poll_times = []

        def fake_hold_state(_ctx, hold_p_id=None):
            poll_times.append(clock.time())
            return {"id": "hold-label", "text": "按住", "display": "none"}

        with (
            patch.object(mod.time, "time", side_effect=clock.time),
            patch.object(mod.time, "sleep", side_effect=clock.sleep),
            patch.object(mod.random, "uniform", side_effect=lambda a, b: (a + b) / 2.0),
            patch.object(mod.random, "randint", return_value=320),
            patch.object(mod, "_resolve_hold_label_for_target", return_value={"id": "hold-label", "text": "Press and hold"}, create=True),
            patch.object(mod, "_px_hold_instruction_state", side_effect=fake_hold_state, create=True),
        ):
            ok = mod._perform_hold("page", ctx, {"x": 100, "y": 200}, 1, 1, "[#1][ruoyi]")

        self.assertAlmostEqual(ok, 5.0, delta=0.01)
        self.assertTrue(poll_times)
        self.assertGreaterEqual(poll_times[0], 105.0)
        self.assertLess(clock.time(), 106.0)
        self.assertIn(("release",), ctx.actions.calls)
        self.assertNotIn(("release_all",), ctx.actions.calls)


    def test_hold_state_returns_exact_press_and_hold_p(self):
        class _Ctx:
            def run_js_loaded(self, _script):
                return {"id": "hidden-label", "text": "Press and hold", "display": "none", "visibility": "visible", "opacity": "1", "hidden": True}

        state = mod._px_hold_instruction_state(_Ctx())
        self.assertEqual("none", (state or {}).get("display"))
        self.assertEqual("hidden-label", (state or {}).get("id"))


    def test_perform_hold_uses_target_hold_p_id_for_early_release_check(self):
        clock = _FakeClock()
        ctx = _FakeContext(clock)
        seen_ids = []

        def fake_hold_state(_ctx, hold_p_id=None):
            seen_ids.append(hold_p_id)
            return {"id": hold_p_id, "text": "Press and hold", "display": "none"}

        with (
            patch.object(mod.time, "time", side_effect=clock.time),
            patch.object(mod.time, "sleep", side_effect=clock.sleep),
            patch.object(mod.random, "uniform", side_effect=lambda a, b: (a + b) / 2.0),
            patch.object(mod.random, "randint", return_value=320),
            patch.object(mod, "_resolve_hold_label_for_target", return_value={"id": "target-p-1", "text": "Press and hold"}, create=True),
            patch.object(mod, "_px_hold_instruction_state", side_effect=fake_hold_state, create=True),
        ):
            ok = mod._perform_hold("page", ctx, {"x": 100, "y": 200}, 1, 1, "[#1][ruoyi]")

        self.assertTrue(ok)
        self.assertEqual(["target-p-1"], seen_ids)
        self.assertLess(clock.time(), 106.0)


    def test_perform_hold_refinds_hold_p_id_when_old_id_disappears(self):
        clock = _FakeClock()
        ctx = _FakeContext(clock)
        seen_ids = []

        def fake_hold_state(_ctx, hold_p_id=None):
            seen_ids.append(hold_p_id)
            if hold_p_id == "new-p-2":
                return {"id": hold_p_id, "text": "Press and hold", "display": "none"}
            return None

        with (
            patch.object(mod.time, "time", side_effect=clock.time),
            patch.object(mod.time, "sleep", side_effect=clock.sleep),
            patch.object(mod.random, "uniform", side_effect=lambda a, b: (a + b) / 2.0),
            patch.object(mod.random, "randint", return_value=320),
            patch.object(mod, "_resolve_hold_label_for_target", side_effect=[{"id": "old-p-1", "text": "Press and hold"}, {"id": "new-p-2", "text": "Press and hold"}], create=True),
            patch.object(mod, "_px_hold_instruction_state", side_effect=fake_hold_state, create=True),
            patch.object(mod, "_find_hold_target", return_value={"x": 110, "y": 210}, create=True),
        ):
            ok = mod._perform_hold("page", ctx, {"x": 100, "y": 200}, 1, 1, "[#1][ruoyi]")

        self.assertTrue(ok)
        self.assertEqual(["old-p-1", "new-p-2"], seen_ids)
        self.assertLess(clock.time(), 106.0)


    def test_find_hold_context_preserves_iframe_box_fallback(self):
        fake_page = object()
        fake_ctx = object()
        with (
            patch.object(mod, "_all_contexts", return_value=[fake_ctx], create=True),
            patch.object(mod, "_find_hold_target", return_value=None, create=True),
            patch.object(mod, "_find_hsprotect_iframe_box", return_value={"text": "iframe-box", "quality": 3}, create=True),
        ):
            ctx, target = mod._find_hold_context(fake_page)

        self.assertIs(ctx, fake_page)
        self.assertEqual("iframe-box", (target or {}).get("text"))

    def test_perform_hold_refinds_press_and_hold_p_from_global_context_after_iframe_box_click(self):
        clock = _FakeClock()
        page_ctx = _FakeContext(clock)
        frame_ctx = object()
        seen = []

        def fake_resolve(ctx, target):
            seen.append(("resolve", ctx, dict(target or {})))
            if ctx is frame_ctx:
                return {"id": "frame-p-7", "text": "Press and hold"}
            return None

        def fake_hold_state(ctx, hold_p_id=None):
            seen.append(("state", ctx, hold_p_id))
            if ctx is frame_ctx and hold_p_id == "frame-p-7":
                return {"id": "frame-p-7", "text": "Press and hold", "display": "none"}
            return None

        with (
            patch.object(mod.time, "time", side_effect=clock.time),
            patch.object(mod.time, "sleep", side_effect=clock.sleep),
            patch.object(mod.random, "uniform", side_effect=lambda a, b: (a + b) / 2.0),
            patch.object(mod.random, "randint", return_value=320),
            patch.object(mod, "_resolve_hold_label_for_target", side_effect=fake_resolve, create=True),
            patch.object(mod, "_px_hold_instruction_state", side_effect=fake_hold_state, create=True),
            patch.object(mod, "_find_hold_context", return_value=(frame_ctx, {"x": 110, "y": 210, "text": "iframe-box"}), create=True),
        ):
            ok = mod._perform_hold("page", page_ctx, {"x": 100, "y": 200, "text": "iframe-box"}, 1, 1, "[#1][ruoyi]")

        self.assertTrue(ok)
        self.assertIn(("state", frame_ctx, "frame-p-7"), seen)
        self.assertLess(clock.time(), 106.0)


if __name__ == "__main__":
    unittest.main()
