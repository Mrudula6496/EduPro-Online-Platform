
import streamlit as st
import pandas as pd
import plotly.express as px

# Page configuration
st.set_page_config(
    page_title="EduPro Analytics",
    page_icon="🎓",
    layout="wide"
)

# Load datasets
@st.cache_data
def load_data():
    courses = pd.read_excel("courses.xlsx")
    students = pd.read_excel("student.xlsx")
    teachers = pd.read_excel("teachers.xlsx")

    return courses, students, teachers

courses, students, teachers = load_data()

# Dashboard title
st.title("🎓 EduPro")
st.subheader("Online Learning Platform Analytics")

# Display dataset information
st.write("### Dataset Information")

col1, col2, col3 = st.columns(3)

with col1:
    st.metric("Total Courses", len(courses))

with col2:
    st.metric("Total Students", len(students))

with col3:
    st.metric("Total Teachers", len(teachers))

# Display data
st.write("### Courses Dataset")
st.dataframe(courses.head())

st.write("### Students Dataset")
st.dataframe(students.head())

st.write("### Teachers Dataset")
st.dataframe(teachers.head())
