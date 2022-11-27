import streamlit as st
from instructions import show_instructions_page
from sked import show_sked_page

page = st.sidebar.selectbox('Menu', ('Sked Software', 'Instructions'))

if page == "Instructions":
    show_instructions_page()
else:
    show_sked_page()