import unittest

import webui.server as server
import webui.scripts as scripts


class WebuiRuoyiResourceBlockFlagTests(unittest.TestCase):
    def test_ruoyi_schema_exposes_block_resources_checkbox(self):
        script = scripts.script_by_id('register_outlook_ruoyi')
        spec = next((arg for arg in script['args'] if arg['flag'] == '--block-resources'), None)
        self.assertIsNotNone(spec)
        self.assertEqual('bool', spec['type'])
        self.assertFalse(spec['default'])

    def test_build_cmd_includes_block_resources_when_checked(self):
        script = scripts.script_by_id('register_outlook_ruoyi')
        cmd = server._build_cmd(script, {'--count': 1, '--block-resources': True})
        self.assertIn('--block-resources', cmd)

    def test_build_cmd_omits_block_resources_when_unchecked(self):
        script = scripts.script_by_id('register_outlook_ruoyi')
        cmd = server._build_cmd(script, {'--count': 1, '--block-resources': False})
        self.assertNotIn('--block-resources', cmd)


if __name__ == '__main__':
    unittest.main()
