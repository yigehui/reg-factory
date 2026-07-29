import unittest
from types import SimpleNamespace
from unittest.mock import patch

import register_outlook_ruoyi as mod


class RuoyiGeoEmulationTests(unittest.TestCase):
    def test_geo_timezone_parser_accepts_nested_timezone_id(self):
        payload = {"timezone": {"id": "America/New_York"}}

        self.assertEqual("America/New_York", mod._geo_timezone_from_payload(payload))

    def test_apply_ruoyi_proxy_geo_emulation_applies_timezone_and_geolocation(self):
        calls = []

        class Emu:
            def set_timezone(self, timezone_id):
                calls.append(("timezone", timezone_id))

            def set_geolocation(self, latitude, longitude, accuracy=100):
                calls.append(("geolocation", latitude, longitude, accuracy))

        page = SimpleNamespace(emulation=Emu())

        with patch.object(
            mod,
            "_probe_proxy_identity",
            return_value={
                "ip": "203.0.113.10",
                "country_code": "US",
                "timezone": "America/New_York",
                "latitude": 40.7128,
                "longitude": -74.0060,
            },
        ):
            ok = mod._apply_ruoyi_proxy_geo_emulation(page, ["127.0.0.1:1080"], "[#1][ruoyi]")

        self.assertTrue(ok)
        self.assertEqual(
            [
                ("timezone", "America/New_York"),
                ("geolocation", 40.7128, -74.0060, 100),
            ],
            calls,
        )

    def test_apply_ruoyi_proxy_geo_emulation_skips_when_geo_is_missing(self):
        page = SimpleNamespace(emulation=SimpleNamespace())

        with patch.object(
            mod,
            "_probe_proxy_identity",
            return_value={"ip": "203.0.113.10", "country_code": "US", "timezone": "", "latitude": None, "longitude": None},
        ):
            ok = mod._apply_ruoyi_proxy_geo_emulation(page, ["127.0.0.1:1080"], "[#1][ruoyi]")

        self.assertFalse(ok)


if __name__ == "__main__":
    unittest.main()
