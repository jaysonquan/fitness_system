from pathlib import Path
from flask import Flask
from models import db, User, Admin, WorkoutRecord, DietRecord, BodyRecord, Food, WorkoutPlan

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "fitness.db"

app = Flask(__name__)
app.config["SQLALCHEMY_DATABASE_URI"] = f"sqlite:///{DB_PATH}"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db.init_app(app)

with app.app_context():
    db.drop_all()
    db.create_all()

    admin = Admin(username="admin", password="123456")
    db.session.add(admin)

    test_user = User(username="test", password="123456", age=20, height=175, weight=70, goal="增肌")
    db.session.add(test_user)
    db.session.flush()

    foods = [
        Food(name="鸡胸肉", calories_per_100g=165),
        Food(name="米饭", calories_per_100g=116),
        Food(name="鸡蛋", calories_per_100g=155),
        Food(name="牛奶", calories_per_100g=54),
        Food(name="香蕉", calories_per_100g=93),
    ]
    db.session.add_all(foods)

    db.session.commit()

    print("数据库初始化完成！")
    print(f"数据库位置：{DB_PATH}")
    print("管理员账号：admin，密码：123456")
    print("测试学生账号：test，密码：123456")