from pathlib import Path
import os
import sqlite3
from flask import Flask, render_template, request, redirect, url_for, session, flash
from datetime import datetime, date, timedelta
from models import db, User, Admin, WorkoutRecord, DietRecord, BodyRecord, Food, WorkoutPlan

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

try:
    from openai import OpenAI
except ImportError:
    OpenAI = None

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "fitness.db"

app = Flask(__name__)

app.config["SQLALCHEMY_DATABASE_URI"] = f"sqlite:///{DB_PATH}"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["TEMPLATES_AUTO_RELOAD"] = True
app.secret_key = "secret-key"

db.init_app(app)

with app.app_context():
    db.create_all()

    if Admin.query.filter_by(username="admin").first() is None:
        admin = Admin(username="admin", password="123456")
        db.session.add(admin)

    if User.query.filter_by(username="test").first() is None:
        test_user = User(
            username="test",
            password="123456",
            age=20,
            height=175,
            weight=70,
            goal="增肌"
        )
        db.session.add(test_user)

    if Food.query.count() == 0:
        foods = [
            Food(name="鸡胸肉", calories_per_100g=165),
            Food(name="米饭", calories_per_100g=116),
            Food(name="鸡蛋", calories_per_100g=155),
            Food(name="牛奶", calories_per_100g=54),
            Food(name="香蕉", calories_per_100g=93),
        ]
        db.session.add_all(foods)

    db.session.commit()


@app.route("/")
def index():
    return redirect(url_for("login"))


@app.route("/ping")
def ping():
    return "Flask 后端运行正常"


@app.route("/debug/db")
def debug_db():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
    tables = [row[0] for row in cursor.fetchall()]

    counts = {}
    for table_name in ["user", "admin", "food", "workout_plan", "workout_record", "diet_record", "body_record"]:
        if table_name in tables:
            cursor.execute(f"SELECT COUNT(*) FROM {table_name}")
            counts[table_name] = cursor.fetchone()[0]
        else:
            counts[table_name] = "表不存在"

    conn.close()

    return {
        "database_path": str(DB_PATH),
        "database_exists": DB_PATH.exists(),
        "tables": tables,
        "table_counts": counts,
        "has_user_table": "user" in tables,
        "has_admin_table": "admin" in tables,
        "has_food_table": "food" in tables,
        "student_test_account": "test / 123456",
        "admin_account": "admin / 123456",
    }


@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        username = request.form["username"]
        password = request.form["password"]

        old_user = User.query.filter_by(username=username).first()
        if old_user:
            flash("用户名已存在")
            return redirect(url_for("register"))

        user = User(username=username, password=password)
        db.session.add(user)
        db.session.commit()
        flash("注册成功，请登录")
        return redirect(url_for("login"))

    return render_template("register.html")


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form["username"]
        password = request.form["password"]

        user = User.query.filter_by(username=username, password=password).first()

        if user:
            session["user_id"] = user.id
            session["username"] = user.username
            return redirect(url_for("dashboard"))

        flash("用户名或密码错误。学生登录请使用 test / 123456；管理员请进入 /admin/login 使用 admin / 123456")

    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/dashboard")
def dashboard():
    if "user_id" not in session:
        return redirect(url_for("login"))

    user = db.session.get(User, session["user_id"])
    workout_count = WorkoutRecord.query.filter_by(user_id=user.id).count()
    diet_count = DietRecord.query.filter_by(user_id=user.id).count()
    plan_count = WorkoutPlan.query.filter_by(user_id=user.id).count()
    completed_plan_count = WorkoutPlan.query.filter_by(user_id=user.id, is_completed=True).count()

    total_burned = db.session.query(db.func.sum(WorkoutRecord.calories)).filter_by(user_id=user.id).scalar() or 0
    total_intake = db.session.query(db.func.sum(DietRecord.calories)).filter_by(user_id=user.id).scalar() or 0
    latest_body = BodyRecord.query.filter_by(user_id=user.id).order_by(BodyRecord.date.desc()).first()

    today = date.today()
    last_7_days = [today - timedelta(days=i) for i in range(6, -1, -1)]
    chart_labels = [day.strftime("%m-%d") for day in last_7_days]

    workout_chart_data = []
    intake_chart_data = []
    burned_chart_data = []

    for day in last_7_days:
        daily_workout_count = WorkoutRecord.query.filter_by(user_id=user.id, date=day).count()
        daily_intake = db.session.query(db.func.sum(DietRecord.calories)).filter_by(user_id=user.id, date=day).scalar() or 0
        daily_burned = db.session.query(db.func.sum(WorkoutRecord.calories)).filter_by(user_id=user.id, date=day).scalar() or 0

        workout_chart_data.append(daily_workout_count)
        intake_chart_data.append(daily_intake)
        burned_chart_data.append(daily_burned)

    body_records = BodyRecord.query.filter_by(user_id=user.id).order_by(BodyRecord.date.asc()).all()
    body_chart_labels = [record.date.strftime("%m-%d") for record in body_records]
    body_chart_weights = [record.weight for record in body_records]

    if plan_count > 0:
        plan_completion_rate = round(completed_plan_count / plan_count * 100, 1)
    else:
        plan_completion_rate = 0

    return render_template(
        "dashboard.html",
        user=user,
        workout_count=workout_count,
        diet_count=diet_count,
        plan_count=plan_count,
        completed_plan_count=completed_plan_count,
        total_burned=total_burned,
        total_intake=total_intake,
        latest_body=latest_body,
        chart_labels=chart_labels,
        workout_chart_data=workout_chart_data,
        intake_chart_data=intake_chart_data,
        burned_chart_data=burned_chart_data,
        body_chart_labels=body_chart_labels,
        body_chart_weights=body_chart_weights,
        plan_completion_rate=plan_completion_rate,
    )


# 智能推荐路由
@app.route("/recommendation")
def recommendation():
    if "user_id" not in session:
        return redirect(url_for("login"))

    user = db.session.get(User, session["user_id"])

    # 基础数据
    workout_count = WorkoutRecord.query.filter_by(user_id=user.id).count()
    total_burned = db.session.query(db.func.sum(WorkoutRecord.calories)).filter_by(user_id=user.id).scalar() or 0
    total_intake = db.session.query(db.func.sum(DietRecord.calories)).filter_by(user_id=user.id).scalar() or 0

    latest_body = BodyRecord.query.filter_by(user_id=user.id).order_by(BodyRecord.date.desc()).first()

    suggestions = []
    videos = []

    # === 规则1：根据目标 ===
    if user.goal == "减脂":
        suggestions.append("建议每周进行 3-5 次有氧训练（跑步、篮球、HIIT）")
        suggestions.append("控制饮食热量，保持摄入 < 消耗")
        videos.append({
            "title": "帕梅拉 10分钟HIIT高强全身燃脂训练",
            "url": "https://player.bilibili.com/player.html?bvid=BV1r7411q7eb"
        })
    elif user.goal == "增肌":
        suggestions.append("建议进行力量训练（卧推、深蹲、硬拉）每周 3-5 次")
        suggestions.append("保证热量盈余 + 高蛋白饮食")
        videos.append({
            "title": "健身新手入门：标准平板卧推教学",
            "url": "https://player.bilibili.com/player.html?bvid=BV1zp4y1t7jD"
        })
        videos.append({
            "title": "从入门到精通：最详细的深蹲教程",
            "url": "https://player.bilibili.com/player.html?bvid=BV1NA411L7Xo"
        })
    elif user.goal == "塑形":
        suggestions.append("建议力量训练 + 有氧结合")
        videos.append({
            "title": "帕梅拉 15分钟全身燃脂训练",
            "url": "https://player.bilibili.com/player.html?bvid=BV1FN4y1z7ny"
        })

    # === 规则2：训练频率 ===
    if workout_count < 3:
        suggestions.append("你的训练频率较低，建议每周至少训练 3 次")

    # === 规则3：热量 ===
    if total_intake > total_burned:
        suggestions.append("当前热量摄入偏高，可以适当减少高热量食物")
    else:
        suggestions.append("当前热量控制良好，可以继续保持")

    # === 规则4：BMI ===
    bmi = None
    if latest_body and latest_body.bmi:
        bmi = latest_body.bmi
        if bmi > 24:
            suggestions.append("BMI 偏高，建议增加有氧运动")
        elif bmi < 18.5:
            suggestions.append("BMI 偏低，建议增加营养摄入")

    return render_template(
        "recommendation.html",
        user=user,
        suggestions=suggestions,
        bmi=bmi,
        total_intake=total_intake,
        total_burned=total_burned,
        videos=videos
    )


# 用户训练质量评分系统
@app.route("/score")
def score():
    if "user_id" not in session:
        return redirect(url_for("login"))

    user = db.session.get(User, session["user_id"])

    workout_count = WorkoutRecord.query.filter_by(user_id=user.id).count()
    diet_count = DietRecord.query.filter_by(user_id=user.id).count()
    body_count = BodyRecord.query.filter_by(user_id=user.id).count()
    plan_count = WorkoutPlan.query.filter_by(user_id=user.id).count()
    completed_plan_count = WorkoutPlan.query.filter_by(user_id=user.id, is_completed=True).count()

    total_intake = db.session.query(db.func.sum(DietRecord.calories)).filter_by(user_id=user.id).scalar() or 0
    total_burned = db.session.query(db.func.sum(WorkoutRecord.calories)).filter_by(user_id=user.id).scalar() or 0

    # 评分维度：训练频率 30分
    workout_score = min(workout_count * 10, 30)

    # 评分维度：饮食记录 20分
    diet_score = min(diet_count * 5, 20)

    # 评分维度：身体指标记录 15分
    body_score = min(body_count * 5, 15)

    # 评分维度：计划完成度 25分
    if plan_count > 0:
        plan_score = round(completed_plan_count / plan_count * 25, 1)
    else:
        plan_score = 0

    # 评分维度：热量管理 10分
    if total_intake == 0 and total_burned == 0:
        calorie_score = 0
    elif user.goal == "减脂" and total_intake <= total_burned:
        calorie_score = 10
    elif user.goal == "增肌" and total_intake >= total_burned:
        calorie_score = 10
    elif user.goal == "塑形":
        calorie_score = 8
    else:
        calorie_score = 5

    total_score = round(workout_score + diet_score + body_score + plan_score + calorie_score, 1)

    if total_score >= 85:
        level = "优秀"
        comment = "训练记录完整，计划执行较好，已具备良好的健身管理习惯。"
    elif total_score >= 70:
        level = "良好"
        comment = "整体表现不错，建议继续提高训练计划完成率和数据记录完整度。"
    elif total_score >= 50:
        level = "一般"
        comment = "目前已有一定记录基础，但训练频率和饮食管理仍需加强。"
    else:
        level = "待提升"
        comment = "当前数据较少，建议先坚持记录训练、饮食和身体指标。"

    score_items = [
        {"name": "训练频率", "score": workout_score, "max": 30},
        {"name": "饮食记录", "score": diet_score, "max": 20},
        {"name": "身体指标", "score": body_score, "max": 15},
        {"name": "计划完成", "score": plan_score, "max": 25},
        {"name": "热量管理", "score": calorie_score, "max": 10},
    ]


    return render_template(
        "score.html",
        user=user,
        total_score=total_score,
        level=level,
        comment=comment,
        score_items=score_items,
        workout_count=workout_count,
        diet_count=diet_count,
        body_count=body_count,
        plan_count=plan_count,
        completed_plan_count=completed_plan_count,
    )


# 简单 AI 推荐系统（规则加权模型）
@app.route("/ai_recommend")
def ai_recommend():
    if "user_id" not in session:
        return redirect(url_for("login"))

    user = db.session.get(User, session["user_id"])

    workout_count = WorkoutRecord.query.filter_by(user_id=user.id).count()
    diet_count = DietRecord.query.filter_by(user_id=user.id).count()
    body_count = BodyRecord.query.filter_by(user_id=user.id).count()
    plan_count = WorkoutPlan.query.filter_by(user_id=user.id).count()
    completed_plan_count = WorkoutPlan.query.filter_by(user_id=user.id, is_completed=True).count()

    total_intake = db.session.query(db.func.sum(DietRecord.calories)).filter_by(user_id=user.id).scalar() or 0
    total_burned = db.session.query(db.func.sum(WorkoutRecord.calories)).filter_by(user_id=user.id).scalar() or 0
    latest_body = BodyRecord.query.filter_by(user_id=user.id).order_by(BodyRecord.date.desc()).first()

    bmi = latest_body.bmi if latest_body and latest_body.bmi else None

    # 模拟机器学习特征工程：将用户行为转化为特征分数
    features = {
        "training_frequency": min(workout_count / 5, 1),
        "diet_tracking": min(diet_count / 7, 1),
        "body_tracking": min(body_count / 3, 1),
        "plan_completion": completed_plan_count / plan_count if plan_count > 0 else 0,
        "calorie_balance": 1 if total_intake >= total_burned else 0.5,
    }

    ai_score = round(
        features["training_frequency"] * 30 +
        features["diet_tracking"] * 20 +
        features["body_tracking"] * 15 +
        features["plan_completion"] * 25 +
        features["calorie_balance"] * 10,
        1
    )

    if user.goal == "增肌":
        model_result = "力量增肌型"
        recommendation_text = "建议采用中高强度力量训练，并保持适度热量盈余。"
        recommended_actions = ["平板卧推", "深蹲", "硬拉", "引体向上"]
    elif user.goal == "减脂":
        model_result = "减脂燃脂型"
        recommendation_text = "建议提高有氧训练频率，并控制每日热量摄入。"
        recommended_actions = ["慢跑", "跳绳", "HIIT", "核心训练"]
    else:
        model_result = "综合塑形型"
        recommendation_text = "建议力量训练和有氧训练结合，提高身体线条和心肺能力。"
        recommended_actions = ["哑铃深蹲", "俯卧撑", "平板支撑", "椭圆机"]

    if bmi:
        if bmi > 24:
            recommendation_text += " 当前 BMI 偏高，建议增加有氧训练比例。"
        elif bmi < 18.5:
            recommendation_text += " 当前 BMI 偏低，建议增加营养摄入和力量训练。"

    deepseek_advice = None
    deepseek_error = None
    api_key = os.getenv("DEEPSEEK_API_KEY")

    if api_key and OpenAI is not None:
        try:
            client = OpenAI(
                api_key=api_key,
                base_url="https://api.deepseek.com"
            )

            prompt = f"""
你是一名专业健身教练和运动健康顾问。请根据以下用户数据，生成个性化健身建议。

用户目标：{user.goal or '未设置'}
训练记录数：{workout_count}
饮食记录数：{diet_count}
身体指标记录数：{body_count}
计划总数：{plan_count}
已完成计划数：{completed_plan_count}
累计摄入热量：{total_intake} kcal
累计运动消耗：{total_burned} kcal
BMI：{bmi if bmi else '暂无'}
本地模型判断：{model_result}
本地模型建议：{recommendation_text}

请用中文输出，结构为：
1. 用户状态判断
2. 训练建议
3. 饮食建议
4. 风险提醒
5. 未来一周执行计划
要求：语言简洁、专业、适合课程设计系统展示，不要输出过长。
"""

            response = client.chat.completions.create(
                model="deepseek-chat",
                messages=[
                    {"role": "system", "content": "你是一个专业、严谨、适合学生健身管理系统使用的AI健身推荐助手。"},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.7,
                max_tokens=800
            )

            deepseek_advice = response.choices[0].message.content
        except Exception as e:
            deepseek_error = str(e)
    elif not api_key:
        deepseek_error = "未检测到 DEEPSEEK_API_KEY，当前展示本地AI推荐结果。"
    elif OpenAI is None:
        deepseek_error = "未安装 openai 依赖，请运行：pip install openai python-dotenv"

    return render_template(
        "ai_recommend.html",
        user=user,
        features=features,
        ai_score=ai_score,
        model_result=model_result,
        recommendation_text=recommendation_text,
        recommended_actions=recommended_actions,
        bmi=bmi,
        deepseek_advice=deepseek_advice,
        deepseek_error=deepseek_error,
    )


@app.route("/profile", methods=["GET", "POST"])
def profile():
    if "user_id" not in session:
        return redirect(url_for("login"))

    user = db.session.get(User, session["user_id"])

    if request.method == "POST":
        user.age = request.form.get("age") or None
        user.height = request.form.get("height") or None
        user.weight = request.form.get("weight") or None
        user.goal = request.form.get("goal") or None
        db.session.commit()
        flash("个人信息已更新")
        return redirect(url_for("profile"))

    return render_template("profile.html", user=user)


@app.route("/workouts")
def workouts():
    if "user_id" not in session:
        return redirect(url_for("login"))

    records = WorkoutRecord.query.filter_by(user_id=session["user_id"]).order_by(WorkoutRecord.date.desc()).all()
    return render_template("workout_records.html", records=records)


@app.route("/workouts/add", methods=["GET", "POST"])
def add_workout():
    if "user_id" not in session:
        return redirect(url_for("login"))

    if request.method == "POST":
        record = WorkoutRecord(
            date=datetime.strptime(request.form["date"], "%Y-%m-%d").date() if request.form.get("date") else date.today(),
            item=request.form["item"],
            duration=request.form.get("duration") or None,
            status=request.form.get("status") or "已完成",
            calories=request.form.get("calories") or 0,
            user_id=session["user_id"]
        )
        db.session.add(record)
        db.session.commit()
        return redirect(url_for("workouts"))

    return render_template("add_workout.html")


@app.route("/diets")
def diets():
    if "user_id" not in session:
        return redirect(url_for("login"))

    records = DietRecord.query.filter_by(user_id=session["user_id"]).order_by(DietRecord.date.desc()).all()
    return render_template("diet_records.html", records=records)


@app.route("/diets/add", methods=["GET", "POST"])
def add_diet():
    if "user_id" not in session:
        return redirect(url_for("login"))

    if request.method == "POST":
        record = DietRecord(
            date=datetime.strptime(request.form["date"], "%Y-%m-%d").date() if request.form.get("date") else date.today(),
            food_name=request.form["food_name"],
            amount=request.form.get("amount") or "",
            calories=request.form.get("calories") or 0,
            user_id=session["user_id"]
        )
        db.session.add(record)
        db.session.commit()
        return redirect(url_for("diets"))

    return render_template("add_diet.html")


@app.route("/bodies")
def bodies():
    if "user_id" not in session:
        return redirect(url_for("login"))

    records = BodyRecord.query.filter_by(user_id=session["user_id"]).order_by(BodyRecord.date.asc()).all()
    labels = [record.date.strftime("%Y-%m-%d") for record in records]
    weights = [record.weight for record in records]
    return render_template("body_records.html", records=records, labels=labels, weights=weights)


@app.route("/bodies/add", methods=["GET", "POST"])
def add_body():
    if "user_id" not in session:
        return redirect(url_for("login"))

    user = db.session.get(User, session["user_id"])

    if request.method == "POST":
        weight = float(request.form["weight"])
        bmi = None
        if user.height:
            bmi = round(weight / ((float(user.height) / 100) ** 2), 2)

        record = BodyRecord(
            date=datetime.strptime(request.form["date"], "%Y-%m-%d").date() if request.form.get("date") else date.today(),
            weight=weight,
            bmi=bmi,
            waist=request.form.get("waist") or None,
            user_id=session["user_id"]
        )
        db.session.add(record)
        db.session.commit()
        return redirect(url_for("bodies"))

    return render_template("add_body.html")


@app.route("/workout_plans")
def workout_plans():
    if "user_id" not in session:
        return redirect(url_for("login"))

    plans = WorkoutPlan.query.filter_by(user_id=session["user_id"]).all()
    return render_template("workout_plans.html", plans=plans)



@app.route("/plans/add", methods=["GET", "POST"])
def add_plan():
    if "user_id" not in session:
        return redirect(url_for("login"))

    if request.method == "POST":
        plan = WorkoutPlan(
            plan_name=request.form["plan_name"],
            action=request.form["action"],
            sets=request.form.get("sets") or None,
            reps=request.form.get("reps") or None,
            weight=request.form.get("weight") or None,
            user_id=session["user_id"]
        )
        db.session.add(plan)
        db.session.commit()
        return redirect(url_for("workout_plans"))

    return render_template("add_plan.html")


# 自动生成训练计划路由
@app.route("/plans/generate")
def generate_plan():
    if "user_id" not in session:
        return redirect(url_for("login"))

    user = db.session.get(User, session["user_id"])
    goal = user.goal or "塑形"

    plan_templates = {
        "增肌": [
            {"plan_name": "增肌训练计划", "action": "平板卧推", "sets": 4, "reps": 8, "weight": 40},
            {"plan_name": "增肌训练计划", "action": "深蹲", "sets": 4, "reps": 10, "weight": 50},
            {"plan_name": "增肌训练计划", "action": "硬拉", "sets": 3, "reps": 6, "weight": 60},
            {"plan_name": "增肌训练计划", "action": "引体向上", "sets": 4, "reps": 6, "weight": 0},
        ],
        "减脂": [
            {"plan_name": "减脂训练计划", "action": "慢跑", "sets": 1, "reps": 30, "weight": 0},
            {"plan_name": "减脂训练计划", "action": "跳绳", "sets": 5, "reps": 100, "weight": 0},
            {"plan_name": "减脂训练计划", "action": "HIIT 波比跳", "sets": 4, "reps": 15, "weight": 0},
            {"plan_name": "减脂训练计划", "action": "核心卷腹", "sets": 4, "reps": 20, "weight": 0},
        ],
        "塑形": [
            {"plan_name": "塑形训练计划", "action": "哑铃深蹲", "sets": 4, "reps": 12, "weight": 20},
            {"plan_name": "塑形训练计划", "action": "俯卧撑", "sets": 4, "reps": 15, "weight": 0},
            {"plan_name": "塑形训练计划", "action": "平板支撑", "sets": 4, "reps": 60, "weight": 0},
            {"plan_name": "塑形训练计划", "action": "椭圆机", "sets": 1, "reps": 25, "weight": 0},
        ],
    }

    selected_plan = plan_templates.get(goal, plan_templates["塑形"])

    for item in selected_plan:
        plan = WorkoutPlan(
            plan_name=item["plan_name"],
            action=item["action"],
            sets=item["sets"],
            reps=item["reps"],
            weight=item["weight"],
            user_id=user.id
        )
        db.session.add(plan)

    db.session.commit()
    flash(f"已根据你的目标“{goal}”自动生成完整训练计划")
    return redirect(url_for("workout_plans"))


@app.route("/plans/edit/<int:plan_id>", methods=["GET", "POST"])
def edit_plan(plan_id):
    if "user_id" not in session:
        return redirect(url_for("login"))

    plan = db.get_or_404(WorkoutPlan, plan_id)
    if plan.user_id != session["user_id"]:
        flash("无权操作")
        return redirect(url_for("workout_plans"))

    if request.method == "POST":
        plan.plan_name = request.form["plan_name"]
        plan.action = request.form["action"]
        plan.sets = request.form.get("sets") or None
        plan.reps = request.form.get("reps") or None
        plan.weight = request.form.get("weight") or None
        db.session.commit()
        return redirect(url_for("workout_plans"))

    return render_template("edit_plan.html", plan=plan)


@app.route("/plans/delete/<int:plan_id>")
def delete_plan(plan_id):
    if "user_id" not in session:
        return redirect(url_for("login"))

    plan = db.get_or_404(WorkoutPlan, plan_id)
    if plan.user_id != session["user_id"]:
        flash("无权操作")
        return redirect(url_for("workout_plans"))

    db.session.delete(plan)
    db.session.commit()
    return redirect(url_for("workout_plans"))


@app.route("/plans/complete/<int:plan_id>")
def complete_plan(plan_id):
    if "user_id" not in session:
        return redirect(url_for("login"))

    plan = db.get_or_404(WorkoutPlan, plan_id)
    if plan.user_id != session["user_id"]:
        flash("无权操作")
        return redirect(url_for("workout_plans"))

    plan.is_completed = not plan.is_completed
    db.session.commit()
    return redirect(url_for("workout_plans"))


@app.route("/admin")
def admin_index():
    return redirect(url_for("admin_login"))


@app.route("/admin/login", methods=["GET", "POST"])
def admin_login():
    if request.method == "POST":
        username = request.form["username"]
        password = request.form["password"]

        admin = Admin.query.filter_by(username=username, password=password).first()
        if admin:
            session["admin_id"] = admin.id
            return redirect(url_for("admin_dashboard"))

        flash("管理员账号或密码错误")

    return render_template("admin_login.html")


@app.route("/admin/dashboard")
def admin_dashboard():
    if "admin_id" not in session:
        return redirect(url_for("admin_login"))

    user_count = User.query.count()
    workout_count = WorkoutRecord.query.count()
    diet_count = DietRecord.query.count()
    food_count = Food.query.count()
    plan_count = WorkoutPlan.query.count()

    return render_template(
        "admin_dashboard.html",
        user_count=user_count,
        workout_count=workout_count,
        diet_count=diet_count,
        food_count=food_count,
        plan_count=plan_count
    )


@app.route("/admin/users")
def admin_users():
    if "admin_id" not in session:
        return redirect(url_for("admin_login"))

    users = User.query.all()
    return render_template("admin_users.html", users=users)


@app.route("/admin/foods", methods=["GET", "POST"])
def admin_foods():
    if "admin_id" not in session:
        return redirect(url_for("admin_login"))

    if request.method == "POST":
        food = Food(
            name=request.form["name"],
            calories_per_100g=request.form["calories_per_100g"]
        )
        db.session.add(food)
        db.session.commit()
        return redirect(url_for("admin_foods"))

    foods = Food.query.all()
    return render_template("admin_foods.html", foods=foods)


@app.route("/admin/foods/delete/<int:food_id>")
def delete_food(food_id):
    if "admin_id" not in session:
        return redirect(url_for("admin_login"))

    food = db.get_or_404(Food, food_id)
    db.session.delete(food)
    db.session.commit()
    return redirect(url_for("admin_foods"))


@app.route("/admin/logout")
def admin_logout():
    session.pop("admin_id", None)
    return redirect(url_for("admin_login"))


if __name__ == "__main__":
    # 再次确保 DB_PATH 一定存在（防止解释器缓存问题）
    from pathlib import Path
    BASE_DIR = Path(__file__).resolve().parent
    DB_PATH = BASE_DIR / "fitness.db"

    print(f"当前使用的数据库：{DB_PATH}")
    print("学生登录地址：http://127.0.0.1:5001/login")
    print("管理员登录地址：http://127.0.0.1:5001/admin/login")
    print("数据库检查地址：http://127.0.0.1:5001/debug/db")
    app.run(debug=True, port=5001)
