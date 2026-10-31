import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# ==========================================================
# EduPro - Online Learning Platform Analytics
# ==========================================================

st.set_page_config(
    page_title="EduPro | Online Learning Analytics",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Theme colours
PURPLE_BG = "#4B328F"
DARK_PURPLE = "#570064"
MID_PURPLE = "#762080"
WHITE = "#FFFFFF"
TEXT = "#30303B"

# ---------------------------- CSS -------------------------
st.markdown(
    f"""
    <style>
    #MainMenu, footer {{visibility: hidden;}}
    header[data-testid="stHeader"] {{background: transparent;}}
    .stApp {{background: {PURPLE_BG};}}
    .block-container {{
        max-width: 100%;
        padding: 0.45rem 0.65rem 0.65rem 0.65rem;
    }}

    /* Default text is dark; only the outer dashboard/header/filter labels are white */
    .stApp {{
        color: #30303B;
    }}
    .stApp p, .stApp h1, .stApp h2, .stApp h3,
    .stApp [data-testid="stMarkdownContainer"] {{
        color: #30303B;
    }}

    /* White visual panels: include the title area as well as the plot */
    div[data-testid="stVerticalBlockBorderWrapper"],
    div[data-testid="stVerticalBlockBorderWrapper"] > div,
    .stVerticalBlockBorderWrapper {{
        background-color: #FFFFFF !important;
        background: #FFFFFF !important;
        border: 1px solid #E5E0EF !important;
        border-radius: 10px !important;
        padding: 7px 8px 4px 8px !important;
        margin-bottom: 8px !important;
        box-shadow: 0 1px 3px rgba(30, 15, 60, 0.10);
    }}
    div[data-testid="stVerticalBlockBorderWrapper"] p,
    div[data-testid="stVerticalBlockBorderWrapper"] label,
    div[data-testid="stVerticalBlockBorderWrapper"] h1,
    div[data-testid="stVerticalBlockBorderWrapper"] h2,
    div[data-testid="stVerticalBlockBorderWrapper"] h3,
    div[data-testid="stVerticalBlockBorderWrapper"] [data-testid="stMarkdownContainer"],
    div[data-testid="stVerticalBlockBorderWrapper"] [data-testid="stMarkdownContainer"] *,
    .stVerticalBlockBorderWrapper p,
    .stVerticalBlockBorderWrapper label,
    .stVerticalBlockBorderWrapper [data-testid="stMarkdownContainer"] {{
        color: #30303B !important;
    }}

    /* Header */
    .edu-header {{
        height: 108px;
        width: 100%;
        box-sizing: border-box;
        display: flex;
        align-items: center;
        padding: 8px 16px;
        color: white !important;
        background: {PURPLE_BG};
        border: 1px solid rgba(25, 10, 65, 0.35);
        margin-bottom: 7px;
    }}
    .edu-cap {{
        font-size: 48px;
        line-height: 1;
        padding-right: 14px;
        color: white !important;
    }}
    .edu-divider {{
        height: 66px;
        width: 2px;
        background: white;
        margin-right: 14px;
        flex: 0 0 2px;
    }}
    .edu-title {{
        color: white !important;
        font-family: Georgia, 'Times New Roman', serif;
        font-size: 39px;
        font-weight: bold;
        line-height: 1.0;
    }}
    .edu-subtitle {{
        color: white !important;
        font-family: Georgia, 'Times New Roman', serif;
        font-size: 18px;
        line-height: 1.2;
        white-space: nowrap;
    }}

    /* Filters */
    div[data-testid="stSelectbox"] {{
        background: white;
        border-radius: 7px;
        padding: 3px 7px 0 7px;
    }}
    div[data-testid="stSelectbox"] label {{
        color: white !important;
        font-size: 12px;
        font-weight: 600;
    }}
    div[data-testid="stSelectbox"] div[data-baseweb="select"] > div {{
        background: white;
        color: #333333;
        border-color: #DDDDDD;
    }}

    /* KPI cards */
    div[data-testid="stMetric"] {{
        background: white;
        border-radius: 8px;
        border-left: 5px solid #282828;
        padding: 9px 8px;
        min-height: 77px;
        text-align: center;
    }}
    div[data-testid="stMetric"] label,
    div[data-testid="stMetric"] [data-testid="stMetricLabel"] {{
        color: #666666 !important;
        justify-content: center;
        font-size: 13px;
    }}
    div[data-testid="stMetric"] [data-testid="stMetricValue"] {{
        color: #292929 !important;
        justify-content: center;
        font-size: 25px;
        font-weight: 600;
    }}
    div[data-testid="stMetric"] *,
    div[data-testid="stMetric"] [data-testid="stMetricLabel"],
    div[data-testid="stMetric"] [data-testid="stMetricValue"],
    div[data-testid="stMetric"] label {{
        color: #292929 !important;
    }}

    .card-title {{
        color: #30303B !important;
        background: #FFFFFF !important;
        text-align: center;
        font-family: Georgia, 'Times New Roman', serif;
        font-size: 13px;
        font-weight: 500;
        margin: 0 0 2px 0;
    }}
    .footer-text {{
        color: white !important;
        text-align: center;
        font-size: 10px;
        padding: 4px;
    }}
    div[data-testid="stPlotlyChart"],
    div[data-testid="stPlotlyChart"] > div {{
        background: #FFFFFF !important;
        border-radius: 7px !important;
    }}
    @media (max-width: 900px) {{
        .edu-title {{font-size: 29px;}}
        .edu-subtitle {{font-size: 14px; white-space: normal;}}
        .edu-cap {{font-size: 38px;}}
        .edu-header {{height: 95px; padding: 6px 9px;}}
    }}
    </style>
    """,
    unsafe_allow_html=True
)

# -------------------------- Data --------------------------
@st.cache_data
def load_data():
    courses = pd.read_excel("courses.xlsx")
    students = pd.read_excel("student.xlsx")
    teachers = pd.read_excel("teachers.xlsx")

    for df in (courses, students, teachers):
        df.columns = df.columns.astype(str).str.strip()

    courses = courses.dropna(how="all").copy()
    students = students.dropna(how="all").copy()
    teachers = teachers.dropna(how="all").copy()

    for col in ["CoursePrice", "CourseDurationHours", "CourseRating"]:
        courses[col] = pd.to_numeric(courses[col], errors="coerce")

    for col in ["Age"]:
        students[col] = pd.to_numeric(students[col], errors="coerce")
        teachers[col] = pd.to_numeric(teachers[col], errors="coerce")

    for col in ["YearsOfExperience", "TeacherRating"]:
        teachers[col] = pd.to_numeric(teachers[col], errors="coerce")

    return courses, students, teachers


try:
    courses, students, teachers = load_data()
except FileNotFoundError as error:
    st.error(
        f"Missing file: {error.filename}. Keep app.py, courses.xlsx, "
        "student.xlsx and teachers.xlsx in the same folder."
    )
    st.stop()
except Exception as error:
    st.error(f"Could not read the Excel files: {error}")
    st.stop()

required_columns = {
    "courses.xlsx": (courses, {
        "CourseID", "CourseName", "CourseCategory", "CourseType",
        "CourseLevel", "CoursePrice", "CourseDurationHours", "CourseRating"
    }),
    "student.xlsx": (students, {"UserID", "UserName", "Age", "Gender", "Email"}),
    "teachers.xlsx": (teachers, {
        "TeacherID", "TeacherName", "Age", "Gender", "Expertise",
        "YearsOfExperience", "TeacherRating"
    })
}
for filename, (dataframe, required) in required_columns.items():
    missing = required - set(dataframe.columns)
    if missing:
        st.error(f"{filename} is missing columns: {', '.join(sorted(missing))}")
        st.stop()

# ----------------------- Chart helpers --------------------
def chart_layout(fig, height=190, legend=True):
    fig.update_layout(
        height=height,
        paper_bgcolor="white",
        plot_bgcolor="white",
        font=dict(family="Arial", size=9, color="#333333"),
        margin=dict(l=10, r=10, t=8, b=28),
        showlegend=legend
    )
    fig.update_xaxes(showgrid=False, zeroline=False, tickfont=dict(size=8))
    fig.update_yaxes(showgrid=True, gridcolor="#EEEEEE", zeroline=False,
                     tickfont=dict(size=8))
    return fig


def draw_card(title, fig, height=190):
    """Render a title and its chart together inside one white card."""
    with st.container(border=True):
        st.markdown(f'<div class="card-title">{title}</div>',
                    unsafe_allow_html=True)
        if fig is None:
            st.caption("No data available for the selected filters.")
        else:
            st.plotly_chart(
                fig, use_container_width=True,
                config={"displayModeBar": False},
                key=f"chart_{title.replace(' ', '_').replace('/', '_')}"
            )


# -------------------------- Header ------------------------
col_header, col_pie, col_level = st.columns([1.38, 0.78, 0.92], gap="small")

with col_header:
    st.markdown(
        """
        <div class="edu-header">
            <div class="edu-cap">🎓</div>
            <div class="edu-divider"></div>
            <div>
                <div class="edu-title">EduPro</div>
                <div class="edu-subtitle">Online Learning Platform Analytics</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    filter1, filter2, filter3, filter4 = st.columns(4, gap="small")
    with filter1:
        course_options = ["All"] + sorted(courses["CourseName"].dropna().astype(str).unique())
        selected_course = st.selectbox("CourseName", course_options)
    with filter2:
        gender_options = ["All"] + sorted(students["Gender"].dropna().astype(str).unique())
        selected_gender = st.selectbox("Gender", gender_options)
    with filter3:
        level_options = ["All"] + sorted(courses["CourseLevel"].dropna().astype(str).unique())
        selected_level = st.selectbox("CourseLevel", level_options)
    with filter4:
        category_options = ["All"] + sorted(courses["CourseCategory"].dropna().astype(str).unique())
        selected_category = st.selectbox("CourseCategory", category_options)

# Apply filters
fc = courses.copy()
fs = students.copy()
ft = teachers.copy()

if selected_course != "All":
    fc = fc[fc["CourseName"] == selected_course]
if selected_level != "All":
    fc = fc[fc["CourseLevel"] == selected_level]
if selected_category != "All":
    fc = fc[fc["CourseCategory"] == selected_category]
if selected_gender != "All":
    fs = fs[fs["Gender"] == selected_gender]
    ft = ft[ft["Gender"] == selected_gender]

# ----------------------- Top pie chart --------------------
with col_pie:
    if fc.empty:
        pie_fig = None
    else:
        pie_data = (
            fc.groupby("CourseCategory", as_index=False)["CourseDurationHours"]
            .sum()
        )
        pie_fig = px.pie(
            pie_data, names="CourseCategory", values="CourseDurationHours",
            color_discrete_sequence=px.colors.qualitative.Bold
        )
        pie_fig.update_traces(
            textposition="outside",
            texttemplate="%{value:.1f}<br>(%{percent})",
            textfont_size=7,
            marker=dict(line=dict(color="white", width=1))
        )
        pie_fig.update_layout(
            height=165, paper_bgcolor="#FFFFFF", plot_bgcolor="#FFFFFF",
            font=dict(family="Arial", size=8, color="#333333"),
            margin=dict(l=2, r=2, t=2, b=35),
            legend=dict(orientation="v", x=1.0, y=0.95, font=dict(size=7))
        )
    draw_card("Course duration and their course category", pie_fig)

# ----------------------- Top level chart ------------------
with col_level:
    if fc.empty:
        level_fig = None
    else:
        level_data = (
            fc.groupby(["CourseLevel", "CourseType"], as_index=False)
            .size().rename(columns={"size": "CourseCount"})
        )
        level_fig = px.bar(
            level_data, x="CourseLevel", y="CourseCount",
            color="CourseType", barmode="stack",
            color_discrete_map={"Free": "#762080", "Paid": "#3F004B"},
            category_orders={"CourseLevel": ["Advanced", "Beginner", "Intermediate"]}
        )
        level_fig = chart_layout(level_fig, 165)
        level_fig.update_layout(
            margin=dict(l=8, r=8, t=5, b=28),
            legend=dict(orientation="h", y=-0.28, x=0.1, font=dict(size=8)),
            xaxis_title="", yaxis_title=""
        )
    draw_card("Courses by Level and Type", level_fig)

# --------------------- Main dashboard ---------------------
left_side, right_side = st.columns([1.02, 1.92], gap="small")

# Left column: treemap and course rating
with left_side:
    if fc.empty:
        tree_fig = None
    else:
        tree_data = fc.dropna(subset=["CourseCategory", "CourseName"]).copy()
        tree_fig = px.treemap(
            tree_data,
            path=["CourseCategory", "CourseName"],
            values="CourseDurationHours",
            color="CourseCategory",
            color_discrete_sequence=[
                "#19A84A", "#E76B32", "#19B8D8", "#7250BF",
                "#690078", "#1930A0", "#DF3F91", "#438CF0",
                "#168FE8", "#D9B600", "#1D777A", "#D84450"
            ]
        )
        tree_fig.update_traces(
            textinfo="label",
            textfont=dict(size=9, color="white"),
            marker=dict(line=dict(color="white", width=1))
        )
        tree_fig.update_layout(
            height=245, paper_bgcolor="white", plot_bgcolor="white",
            margin=dict(l=2, r=2, t=2, b=2), showlegend=False
        )
    draw_card("Course Categories", tree_fig)

    if fc.empty:
        rating_fig = None
    else:
        rating_data = (
            fc.dropna(subset=["CourseRating", "CourseName"])
            .nlargest(5, "CourseRating")
            .sort_values("CourseRating", ascending=True)
        )
        rating_fig = px.bar(
            rating_data, x="CourseRating", y="CourseName",
            orientation="h", text="CourseRating",
            color_discrete_sequence=[DARK_PURPLE]
        )
        rating_fig.update_traces(
            texttemplate="%{text:.2f}", textposition="inside",
            textfont=dict(color="white", size=9)
        )
        rating_fig.update_layout(
            height=185, paper_bgcolor="white", plot_bgcolor="white",
            margin=dict(l=5, r=8, t=5, b=22), showlegend=False,
            xaxis=dict(range=[0, 5], title=""),
            yaxis=dict(title="", tickfont=dict(size=8))
        )
    draw_card("Course Rating", rating_fig)

# Right side: KPI cards and charts
with right_side:
    kpi1, kpi2, kpi3 = st.columns(3, gap="small")
    with kpi1:
        with st.container(border=True):
            st.metric("total_students", f"{len(fs):,}")
    with kpi2:
        with st.container(border=True):
            st.metric("Total_experties", f"{ft['Expertise'].count():,}")
    with kpi3:
        with st.container(border=True):
            st.metric("Total_teachers", f"{len(ft):,}")

    free_col, teacher_col = st.columns([0.86, 1.55], gap="small")

    with free_col:
        if fc.empty:
            free_fig = None
        else:
            free_data = (
                fc["CourseType"].value_counts()
                .rename_axis("CourseType").reset_index(name="Count")
            )
            free_fig = px.pie(
                free_data, names="CourseType", values="Count", hole=0.68,
                color="CourseType",
                color_discrete_map={"Free": "#BDBDBD", "Paid": DARK_PURPLE}
            )
            free_fig.update_traces(
                textposition="outside",
                texttemplate="%{value} (%{percent})",
                textfont_size=8,
                marker=dict(line=dict(color="white", width=1))
            )
            free_fig.update_layout(
                height=175, paper_bgcolor="white", plot_bgcolor="white",
                margin=dict(l=5, r=5, t=4, b=22),
                legend=dict(orientation="v", x=1.0, y=0.65, font=dict(size=8))
            )
        draw_card("Free VS Paid", free_fig)

    with teacher_col:
        teacher_top = (
            ft.dropna(subset=["TeacherRating", "TeacherName"])
            .nlargest(10, "TeacherRating")
            .sort_values("TeacherRating", ascending=False)
        )
        if teacher_top.empty:
            teacher_fig = None
        else:
            teacher_fig = px.bar(
                teacher_top, x="TeacherName", y="TeacherRating",
                color_discrete_sequence=[DARK_PURPLE]
            )
            teacher_fig.update_layout(
                height=175, paper_bgcolor="white", plot_bgcolor="white",
                margin=dict(l=5, r=5, t=5, b=42), showlegend=False,
                yaxis=dict(range=[0, 5], title=""),
                xaxis=dict(title="", tickfont=dict(size=7))
            )
        draw_card("Teacher Rating", teacher_fig)

    exp_col, price_col, duration_col = st.columns([1.05, 1.05, 1.0], gap="small")

    with exp_col:
        exp_data = (
            ft.dropna(subset=["YearsOfExperience", "TeacherRating"])
            .groupby("YearsOfExperience", as_index=False)["TeacherRating"]
            .mean().sort_values("YearsOfExperience")
        )
        if exp_data.empty:
            exp_fig = None
        else:
            exp_fig = px.area(
                exp_data, x="YearsOfExperience", y="TeacherRating",
                color_discrete_sequence=["#8C4B91"]
            )
            exp_fig.update_traces(
                line=dict(color=DARK_PURPLE, width=1.5),
                fillcolor="rgba(108, 0, 112, 0.65)"
            )
            exp_fig.update_layout(
                height=175, paper_bgcolor="white", plot_bgcolor="white",
                margin=dict(l=8, r=5, t=5, b=25), showlegend=False,
                xaxis_title="", yaxis_title="", yaxis=dict(range=[0, 5])
            )
        draw_card("Teacher Rating by Years Of Experience", exp_fig)

    with price_col:
        price_data = (
            fc.dropna(subset=["CoursePrice", "CourseName"])
            .sort_values("CoursePrice", ascending=False)
            .head(5).sort_values("CoursePrice", ascending=False)
        )
        if price_data.empty:
            price_fig = None
        else:
            price_fig = px.bar(
                price_data, x="CourseName", y="CoursePrice",
                color_discrete_sequence=[DARK_PURPLE]
            )
            price_fig.update_layout(
                height=175, paper_bgcolor="white", plot_bgcolor="white",
                margin=dict(l=5, r=5, t=5, b=48), showlegend=False,
                xaxis_title="", yaxis_title="",
                xaxis=dict(tickfont=dict(size=7))
            )
        draw_card("CoursePrice by CourseName", price_fig)

    with duration_col:
        total_duration = float(fc["CourseDurationHours"].fillna(0).sum())
        max_duration = max(total_duration * 2, 1)
        duration_fig = go.Figure(go.Indicator(
            mode="gauge+number",
            value=total_duration,
            number={"valueformat": ".2s", "font": {"size": 27, "color": "#777777"}},
            gauge={
                "shape": "angular",
                "axis": {
                    "range": [0, max_duration],
                    "tickmode": "array",
                    "tickvals": [0, max_duration],
                    "ticktext": ["0", f"{max_duration:,.0f}"],
                    "tickfont": {"size": 8, "color": "#777777"}
                },
                "bar": {"color": DARK_PURPLE, "thickness": 0.32},
                "bgcolor": "white",
                "borderwidth": 0,
                "steps": [{"range": [0, max_duration], "color": "#D0D0D0"}]
            }
        ))
        duration_fig.update_layout(
            height=175, paper_bgcolor="white",
            margin=dict(l=8, r=8, t=5, b=2)
        )
        draw_card("Course Duration", duration_fig)

# -------------------------- Footer ------------------------
st.markdown(
    '<div class="footer-text">EduPro | Online Learning Platform Analytics</div>',
    unsafe_allow_html=True
)
