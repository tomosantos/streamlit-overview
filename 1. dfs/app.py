import streamlit as st
import pandas as pd

st.title("Streamlit Elements Demo")

# DataFrame Section
st.subheader("Dataframe")
df = pd.DataFrame({
    'Name': ['Alice', 'Bob', 'Charlie', 'David'],
    'Age': [25, 32, 37, 47],
    'Occupation': ['Engineer', 'Doctor', 'Artist', 'Chef']
})

st.dataframe(df)

# Data Editor Section (editable Dataframe)
st.subheader('Data Editor')
editable_df = st.data_editor(df)

# Static Table Section
st.subheader('Static Table')
st.table(df)

# Metrics Section
st.subheader('Metrics')
st.metric(label="Total Rows", value=len(df))
st.metric(label='Average Age', value=round(df['Age'].mean(), 1))

# JSON and Dict Section
st.subheader("JSON and Dictionary")
sample_dict = {
    'name': 'Alice',
    'age': 25,
    'skills': ['Python', 'Data Science', 'Machine Learning']
}
st.json(sample_dict)


# Also show it as dictionary
st.write('Dictionary view: ', sample_dict)