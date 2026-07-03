import unittest


class RestoredFitnessAppRoutesTest(unittest.TestCase):
    def setUp(self):
        import app as app_module
        from models import Admin, Food, User, db

        self.app_module = app_module
        self.app = app_module.app
        self.client = self.app.test_client()
        self.db = db
        self.User = User
        self.Admin = Admin
        self.Food = Food

        with self.app.app_context():
            db.drop_all()
            db.create_all()
            db.session.add(Admin(username="admin", password="123456"))
            db.session.add(
                User(
                    username="test",
                    password="123456",
                    age=20,
                    height=175,
                    weight=70,
                    goal="增肌",
                )
            )
            db.session.add(Food(name="鸡胸肉", calories_per_100g=165))
            db.session.commit()

    def test_home_redirects_to_student_login(self):
        response = self.client.get("/")

        self.assertEqual(response.status_code, 302)
        self.assertIn("/login", response.location)

    def test_student_can_login_and_open_dashboard(self):
        response = self.client.post(
            "/login",
            data={"username": "test", "password": "123456"},
            follow_redirects=True,
        )

        body = response.get_data(as_text=True)
        self.assertEqual(response.status_code, 200)
        self.assertIn("欢迎你，test", body)
        self.assertIn("最近7天训练次数", body)

    def test_deepseek_recommendation_page_uses_local_fallback_without_api_key(self):
        self.client.post(
            "/login",
            data={"username": "test", "password": "123456"},
            follow_redirects=True,
        )

        response = self.client.get("/ai_recommend")
        body = response.get_data(as_text=True)

        self.assertEqual(response.status_code, 200)
        self.assertIn("DeepSeek AI智能分析", body)
        self.assertIn("未检测到 DEEPSEEK_API_KEY", body)

    def test_student_can_generate_workout_plan(self):
        self.client.post(
            "/login",
            data={"username": "test", "password": "123456"},
            follow_redirects=True,
        )

        response = self.client.get("/plans/generate", follow_redirects=True)
        body = response.get_data(as_text=True)

        self.assertEqual(response.status_code, 200)
        self.assertIn("增肌训练计划", body)
        self.assertIn("平板卧推", body)

    def test_admin_can_manage_foods(self):
        login = self.client.post(
            "/admin/login",
            data={"username": "admin", "password": "123456"},
            follow_redirects=True,
        )
        self.assertEqual(login.status_code, 200)
        self.assertIn("管理员后台总览", login.get_data(as_text=True))

        response = self.client.post(
            "/admin/foods",
            data={"name": "燕麦", "calories_per_100g": "380"},
            follow_redirects=True,
        )
        body = response.get_data(as_text=True)

        self.assertEqual(response.status_code, 200)
        self.assertIn("燕麦", body)


if __name__ == "__main__":
    unittest.main()
