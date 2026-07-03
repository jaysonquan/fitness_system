import unittest


class DeepSeekRestoreTest(unittest.TestCase):
    def test_ai_recommend_route_is_restored(self):
        import app as app_module

        flask_app = app_module.app
        routes = {rule.rule for rule in flask_app.url_map.iter_rules()}

        self.assertIn("/ai_recommend", routes)


if __name__ == "__main__":
    unittest.main()
